package geoQueryFiller;

import java.io.File;
import org.apache.poi.ss.usermodel.Workbook;

import com.google.appengine.labs.repackaged.com.google.common.base.Objects;
import com.google.code.geocoder.model.GeocoderAddressComponent;
import com.google.code.geocoder.model.GeocoderResult;

import utility.AddressComponent;
import utility.GoogleGeocodeQuerier;
import utility.XLSUtil;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelGeoQueryFiller {

	final static String COLUMN_LOCATION = "Location";
	final static String COLUMN_IS_PROCESSED = "Is Processed";
	final static String COLUMN_IS_FULL_MATCH = "Is Full Match";
	final static String COLUMN_HAS_RESULT = "Has Result";
	final static String COLUMN_FORMATTED_ADDRESS = "Formatted Address";
	final static String COLUMN_SEC_RESULT_START = "Sec. Start";
	final static String COLUMN_SEC_RESULT_END = "Sec. End";

	final static String FULL_GOOGLE_NAME = "Google name";

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		File inputFile = new File(args.length == 0 ? "" : args[0]);
		if (args.length == 0)
			System.out.println("Usage: [XLS/X source file]");
		else if (!inputFile.exists())
			System.out.println("File " + inputFile.getName() + " could not be found.");
		else if (!inputFile.canWrite() || !inputFile.canRead())
			System.out.println("File " + inputFile.getName() + " is not writable.");
		else {
			processFile(inputFile);
		}

	}

	private static void processFile(File inputFile) {
		Workbook wb = XLSUtil.openXLS(inputFile);

		int mainSheet = 0;
		int secSheet = XLSUtil.getOrCreateSheet(wb, "Secondary", 1);
		if (secSheet == 0)
			secSheet = XLSUtil.getOrCreateSheet(wb, "Secondary 2", 1);

		// Find the header row
		int mainRow = -1;
		Map<String, Integer> mainHeader = null;
		boolean isValidHeader = false;
		do {
			++mainRow;
			mainHeader = XLSUtil.getHeader(wb, mainSheet, mainRow);
			isValidHeader = mainHeader.get(COLUMN_LOCATION) != null && mainHeader.get(COLUMN_IS_PROCESSED) != null;
		} while ((isValidHeader == false) && (XLSUtil.isEndRow(wb, mainSheet, mainRow) == false));
		int mainHeaderRow = mainRow++;

		if (isValidHeader == false) {
			System.out.println("No header with rows \"" + COLUMN_LOCATION + "\" and \"" + COLUMN_IS_PROCESSED +
					"\" found in main sheet.");
		} else {
			// Find the header row in the secondary sheet
			int secRow = 0;
			Map<String, Integer> secHeader = XLSUtil.getHeader(wb, secSheet, secRow);
			int secHeaderRow = secRow++;

			// Find the last row in the secondary sheet
			while (XLSUtil.isEndRow(wb, secSheet, secRow) == false)
				++secRow;

			GoogleGeocodeQuerier geoQuerier = new GoogleGeocodeQuerier("en");
			geoQuerier.setRemoveComponent(Arrays.asList(AddressComponent.acCountry, AddressComponent.acStreetAddress,
					AddressComponent.acRoute, AddressComponent.acIntersection, AddressComponent.acPremise,
					AddressComponent.acSubpremise, AddressComponent.acAirport, AddressComponent.acPark));

			++mainRow;
			boolean failedQuery = false;
			while (!XLSUtil.isEndRow(wb, mainSheet, mainRow) && failedQuery == false) {
				String location = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LOCATION));
				String isProcessed = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_IS_PROCESSED));
				// Check if row needs to be written
				if ((location.isEmpty() == false) && (Boolean.valueOf(isProcessed) == false)) {
					List<GeocoderResult> queryRes = geoQuerier.makeQuery(location);
					if (queryRes == null)
						failedQuery = true;
					else {
						// Write the query result
						boolean hasResult = !queryRes.isEmpty();
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_IS_PROCESSED, "True");
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_HAS_RESULT,
								Boolean.toString(hasResult));
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_SEC_RESULT_START,
								(queryRes.size() <= 1) ? "" : Integer.toString(secRow + 1));
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_SEC_RESULT_END,
								(queryRes.size() <= 1) ? "" : Integer.toString(secRow + queryRes.size()));
						writeGeoResultToRow(wb, mainRow, mainSheet, mainHeader, hasResult ? queryRes.get(0) : null);
						++mainRow;

						// Write the non-main result
						for (int i = 1; i < queryRes.size(); ++i) {
							writeGeoResultToRow(wb, secRow, secSheet, secHeader, queryRes.get(i));
							++secRow;
						}
					}
				}
			}

			// Update the headers
			XLSUtil.updateHeader(wb, mainSheet, mainHeaderRow, mainHeader);
			XLSUtil.updateHeader(wb, secSheet, secHeaderRow, secHeader);

			// save the xls
			XLSUtil.saveXLS(wb, inputFile);
		}
	}

	/**
	 * @param wb
	 * @param readRow
	 * @param sheetNum
	 * @param workHeader
	 * @param mainResult
	 */
	private static void writeGeoResultToRow(Workbook wb, int readRow, int sheetNum, Map<String, Integer> workHeader,
			GeocoderResult mainResult) {
		if (mainResult != null) {
			String typeString = "";
			for (String type : mainResult.getTypes()) {
				if (typeString.isEmpty())
					typeString = type;
				else
					typeString = typeString + ", " + type;
			}
			setCellByHeader(wb, sheetNum, readRow, workHeader, "Formatted Address", mainResult.getFormattedAddress());
			setCellByHeader(wb, sheetNum, readRow, workHeader, "Type", typeString);
			setCellByHeader(wb, sheetNum, readRow, workHeader, "Latitude", mainResult.getGeometry().getLocation()
					.getLat().toString());
			setCellByHeader(wb, sheetNum, readRow, workHeader, "Longitude", mainResult.getGeometry().getLocation()
					.getLng().toString());
			for (GeocoderAddressComponent comp : mainResult.getAddressComponents()) {
				for (String type : comp.getTypes()) {
					setCellByHeader(wb, sheetNum, readRow, workHeader, "GT " + type, comp.getLongName());
				}
			}
		}
	}

	private static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, String value) {
		Integer columnIndex = workHeader.get(columnName);
		if (columnIndex == null) {
			columnIndex = new Integer(0);
			for (Integer occupiedColumn : workHeader.values()) {
				columnIndex = Math.max(occupiedColumn + 1, columnIndex);
			}
			workHeader.put(columnName, columnIndex);
		}

		XLSUtil.setCellString(wb, sheetNum, rowNum, columnIndex, value);
	}

}
