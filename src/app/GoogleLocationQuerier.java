package app;

import java.io.File;
import java.io.PrintStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.util.Objects;
import com.google.code.geocoder.model.GeocoderAddressComponent;
import com.google.code.geocoder.model.GeocoderResult;
import com.google.code.geocoder.model.LatLng;
import com.pairapp.engine.parser.data.VariantTypeEnums.CountryState;

import utility.AddressComponent;
import utility.GoogleGeocodeQuerier;
import utility.XLSUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;
import java.util.TreeMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

public class GoogleLocationQuerier {

	enum MldEntryType {
		Location,
		General;
		
		static public MldEntryType toEnum(String value)
		{
			if (value != null)
			{
				for(MldEntryType type : MldEntryType.values())
				{
					if (type.name().equalsIgnoreCase(value)) 
						return type;
				}
			}
			return null;
		}
	}

	private final static String COLUMN_LOCATION = "Location";
	private final static String COLUMN_IS_PROCESSED = "Is Processed";
	private final static String COLUMN_HAS_RESULT = "Has Result";
	private final static String COLUMN_FORMATTED_ADDRESS = "Formatted Address";
	private final static String COLUMN_SEC_RESULT_START = "Sec. Start";
	private final static String COLUMN_SEC_RESULT_END = "Sec. End";
	private final static String COLUMN_ORIG_ROW = "Orig. Row";
	private final static String COLUMN_TYPE = "Type";
	private final static String COLUMN_LATITUDE = "Latitude";
	private final static String COLUMN_LONGITUDE = "Longitude";
	private final static String COLUMN_GOOGLE_TYPE_PREFIX = "GT ";
	
	
	
	final static String FULL_GOOGLE_NAME = "SearchName";

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		
		processFile(new File("E:\\MyProjects\\WispaResources\\Server\\Parser\\Location Data\\dc.xlsx"));
		/*{
			File inputFile = new File(args.length == 0 ? "" : args[0]);
			if (args.length == 0)
				System.out.println("Usage: [XLS/X source file]");
			else if (!inputFile.exists())
				System.out.println("File " + inputFile.getName() + " could not be found.");
			else if (!inputFile.canWrite() || !inputFile.canRead())
				System.out.println("File " + inputFile.getName() + " is not writable.");
			else {
				processFile(inputFile);
				System.out.println("Work Finished.");
			}
		}*/
	}

	private static void processFile(File inputFile) {
		Workbook wb = XLSUtil.openXLS(inputFile);

		int mainSheet = 0;
		
		// Find the header row
		Map<String, Integer> mainHeader = new TreeMap<String, Integer>(String.CASE_INSENSITIVE_ORDER);
		int mainHeaderRow = getHeaderRow(wb, mainSheet, mainHeader, COLUMN_LOCATION);
		if (mainHeaderRow == -1) {
			System.out.println("No header with rows \"" + COLUMN_LOCATION + "\" and \"" + COLUMN_IS_PROCESSED +
					"\" found in main sheet.");
		} else {
			
			addColumnToHeader(mainHeader, COLUMN_ORIG_ROW);
			addColumnToHeader(mainHeader, COLUMN_TYPE);
			addColumnToHeader(mainHeader, COLUMN_LONGITUDE);
			addColumnToHeader(mainHeader, COLUMN_LATITUDE);
			addColumnToHeader(mainHeader, COLUMN_IS_PROCESSED);
			
			
			boolean hasUpdated = updateGeoDataInExcel(wb, mainSheet, mainHeader, mainHeaderRow);
			if (hasUpdated)
			{
				XLSUtil.saveXLS(wb, inputFile);
				return;
			}
		}
	}

	/**
	 * @param wb
	 * @param mainSheet
	 * @param secSheet
	 * @param mainRow
	 * @param mainHeader
	 * @param mainHeaderRow
	 * @param calculatedLoc
	 */
	private static boolean updateGeoDataInExcel(Workbook wb, int mainSheet, 
			Map<String, Integer> mainHeader, int mainHeaderRow) {
		boolean hasUpdated = false;
		//Find the secondary sheet
		int secSheet = XLSUtil.getOrCreateSheet(wb, "Secondary", 1);
		if (secSheet == 0)
			secSheet = XLSUtil.getOrCreateSheet(wb, "Secondary 2", 1);

		// Find the header row in the secondary sheet
		int secRow = 0;
		Map<String, Integer> secHeader = XLSUtil.getHeader(wb, secSheet, secRow);
		int secHeaderRow = secRow++;

		
		//Get the locations that already have been calculated
		Map<String, Integer> calculatedLoc = getPreCalculatedLocations(wb, mainSheet, mainHeaderRow, mainHeader);
		
		// Find the last row in the secondary sheet
		while (XLSUtil.isEndRow(wb, secSheet, secRow) == false)
			++secRow;
		
		GoogleGeocodeQuerier geoQuerier = new GoogleGeocodeQuerier("en");
		//geoQuerier.setRemoveComponent(Arrays.asList(AddressComponent.acCountry, AddressComponent.acStreetAddress,
		//		AddressComponent.acRoute, AddressComponent.acIntersection, AddressComponent.acPremise,
		//		AddressComponent.acSubpremise, AddressComponent.acAirport, AddressComponent.acPark, 
		//		AddressComponent.acSynagogue, AddressComponent.acChurch));
		geoQuerier.setOrderByComponent(Arrays.asList(AddressComponent.acSubwayStation, AddressComponent.acTrainStation, AddressComponent.acTransitStation));
		//subway_station, transit_station, train_station, establishment
		boolean failedQuery = false;
		//Scan all locations query them and output the results
		int mainRow = mainHeaderRow + 1;
		while (!XLSUtil.isEndRow(wb, mainSheet, mainRow) && failedQuery == false) {
			String location = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LOCATION));
			String isProcessed = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_IS_PROCESSED));
			// Check if row needs to be written
			if ((calculatedLoc.get(location) != null) && ((isProcessed.isEmpty() || isProcessed.equalsIgnoreCase("false"))))
				setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_IS_PROCESSED, "Already calculated (" + Integer.toString(calculatedLoc.get(location) + 1) + ")");
			else if ((location.isEmpty() == false) && ((isProcessed.isEmpty() || isProcessed.equalsIgnoreCase("false")))) {
				List<GeocoderResult> queryRes = geoQuerier.makeQuery(location);
				if (queryRes == null)
				{
					System.out.println("Quit due to failed query.");
					failedQuery = true;
				}
				else {
					//Find the best fit
					int fitIndex = 0;
					String trimmedLocation = location.split(",")[0].trim();
					for(int i = 0 ; i < queryRes.size(); ++i)
					{
						GeocoderResult res = queryRes.get(i);
						String trimmedRes = res.getFormattedAddress().split(",")[0].trim();
						if (trimmedRes.equalsIgnoreCase(trimmedLocation))
						{
							fitIndex = i;
							break;							
						}
					}
					
					setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_IS_PROCESSED, true);
					setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_HAS_RESULT, fitIndex != -1);
					
					if (fitIndex != -1)
					{
						writeGeoResultToRow(wb, mainSheet, mainRow, mainHeader, queryRes.get(fitIndex), true);
						queryRes.remove(fitIndex);
					}
					
					if (queryRes.size() > 0)
					{
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_SEC_RESULT_START, secRow + 1);
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_SEC_RESULT_END, secRow + queryRes.size() + 1);
					}
					

					// Write the non-main result
					for (int i = 0; i < queryRes.size(); ++i) {
						writeGeoResultToRow(wb, secSheet, secRow, secHeader, queryRes.get(i), false);
						++secRow;
					}
					calculatedLoc.put(location, mainRow);
					hasUpdated = true;
				}
			}
			++mainRow;
		}

		// Update the headers
		XLSUtil.updateHeader(wb, mainSheet, mainHeaderRow, mainHeader);
		XLSUtil.updateHeader(wb, secSheet, secHeaderRow, secHeader);
		return hasUpdated;
	}

	private static Map<String, Integer> getPreCalculatedLocations(Workbook wb, int mainSheet, int mainHeaderRow,
			Map<String, Integer> mainHeader) {
		Map<String, Integer> precalced = new TreeMap<String, Integer>(String.CASE_INSENSITIVE_ORDER);
		
		int mainRow = mainHeaderRow + 1;
		//Scan all locations query them and output the results
		while (!XLSUtil.isEndRow(wb, mainSheet, mainRow)) {
			String location = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LOCATION));
			String isProcessed = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_IS_PROCESSED));
			// Check if row needs to be written
			if (Boolean.valueOf(isProcessed) == true)
				precalced.put(location, mainRow);
			++mainRow;
		}
		return precalced;
	}

	/**
	 * @param wb
	 * @param readRow
	 * @param sheetNum
	 * @param workHeader
	 * @param mainResult
	 */
	private static void writeGeoResultToRow(Workbook wb, int sheetNum, int readRow, Map<String, Integer> workHeader,
			GeocoderResult mainResult, boolean isMainSheet) {
		if (mainResult != null) {
			String typeString = "";
			for (String type : mainResult.getTypes()) {
				if (typeString.isEmpty())
					typeString = type;
				else
					typeString = typeString + ", " + type;
			}
			if (!isMainSheet)
			{
				setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_ORIG_ROW, Integer.toString(readRow + 1));
			}
			setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_FORMATTED_ADDRESS, mainResult.getFormattedAddress());
			setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_TYPE, typeString);
			setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_LATITUDE, mainResult.getGeometry().getLocation()
					.getLat().doubleValue());
			setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_LONGITUDE, mainResult.getGeometry().getLocation()
					.getLng().doubleValue());
			for (GeocoderAddressComponent comp : mainResult.getAddressComponents()) {
				for (String type : comp.getTypes()) {
					setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_GOOGLE_TYPE_PREFIX + type, comp.getLongName());
				}
			}
		}
	}

	private static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, String value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellString(wb, sheetNum, rowNum, columnIndex, value);
	}
	
	private static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, Boolean value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}
	private static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, Integer value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}
	private static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, Double value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}

	/**
	 * @param workHeader
	 * @param columnName
	 * @return
	 */
	private static Integer addColumnToHeader(Map<String, Integer> workHeader, String columnName) {
		Integer columnIndex = workHeader.get(columnName);
		if (columnIndex == null) {
			columnIndex = new Integer(0);
			for (Integer occupiedColumn : workHeader.values()) {
				columnIndex = Math.max(occupiedColumn + 1, columnIndex);
			}
			workHeader.put(columnName, columnIndex);
		}
		return columnIndex;
	}
	
	private static int getHeaderRow(Workbook wb, int sheetNum, Map<String, Integer> headerColumns, String ... neededColumns)
	{
		int rowNum = -1;
		boolean isValidHeader = false;
		do {
			++rowNum;
			Map<String, Integer> tmpColumns = XLSUtil.getHeader(wb, sheetNum, rowNum);
			isValidHeader = true;
			for(String neededColumn : neededColumns)
			{	
				if ((neededColumn != null) && (neededColumn.isEmpty() == false))
				{	
					isValidHeader &= tmpColumns.get(neededColumn) != null;
					if (isValidHeader)
						headerColumns.putAll(tmpColumns);
				}
			}
		} while ((isValidHeader == false) && (XLSUtil.isEndRow(wb, sheetNum, rowNum) == false));
		
		if (isValidHeader == true)
			return rowNum;
		else return -1;
	}
}
