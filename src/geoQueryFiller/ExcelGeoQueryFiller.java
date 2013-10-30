package geoQueryFiller;

import java.io.File;
import java.io.PrintStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.google.code.geocoder.model.GeocoderAddressComponent;
import com.google.code.geocoder.model.GeocoderResult;

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

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

public class ExcelGeoQueryFiller {

	final static String COLUMN_LOCATION = "Location";
	final static String COLUMN_IS_PROCESSED = "Is Processed";
	final static String COLUMN_IS_FULL_MATCH = "Is Full Match";
	final static String COLUMN_HAS_RESULT = "Has Result";
	final static String COLUMN_FORMATTED_ADDRESS = "Formatted Address";
	final static String COLUMN_SEC_RESULT_START = "Sec. Start";
	final static String COLUMN_SEC_RESULT_END = "Sec. End";
	final static String COLUMN_ORIG_ROW = "Orig. Row";
	final static String COLUMN_TYPE = "Type";
	final static String COLUMN_LATITUDE = "Latitude";
	final static String COLUMN_LONGITUDE = "Longitude";
	final static String COLUMN_GOOGLE_TYPE_PREFIX = "GT ";
	final static String COLUMN_ALTERNATIVE_NAME = "Alternative Names";
	final static String COLUMN_GROUP = "Group";
	
	final static String COLUMN_XML_LEVEL = "Level";
	final static String COLUMN_XML_NAME = "Name";
	final static String COLUMN_XML_FROM = "From";
	

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
			System.out.println("Work Finished.");
		}

	}

	private static void processFile(File inputFile) {
		Workbook wb = XLSUtil.openXLS(inputFile);

		int mainSheet = 0;
		
		// Find the header row
		Map<String, Integer> mainHeader = new HashMap<String, Integer>();
		int mainHeaderRow = getHeaderRow(wb, mainSheet, mainHeader, COLUMN_LOCATION, COLUMN_IS_PROCESSED);
		int mainRow = mainHeaderRow + 1;
		
		if (mainHeaderRow == -1) {
			System.out.println("No header with rows \"" + COLUMN_LOCATION + "\" and \"" + COLUMN_IS_PROCESSED +
					"\" found in main sheet.");
		} else {
			
			addColumnToHeader(mainHeader, COLUMN_ORIG_ROW);
			addColumnToHeader(mainHeader, COLUMN_TYPE);
			addColumnToHeader(mainHeader, COLUMN_LATITUDE);
			addColumnToHeader(mainHeader, COLUMN_LONGITUDE);
			addColumnToHeader(mainHeader, COLUMN_ALTERNATIVE_NAME);
			addColumnToHeader(mainHeader, COLUMN_GROUP);

			HashMap<String, Integer> calculatedLoc = getPreCalculatedLocations(wb, mainSheet, mainRow, mainHeader);
			
			boolean hasUpdated = updateGeoDataInExcel(wb, mainSheet, mainRow, mainHeader, mainHeaderRow, calculatedLoc);
			// save the xls
			if (hasUpdated)
				XLSUtil.saveXLS(wb, inputFile);
			
			
			exportToXML(wb, mainSheet, mainRow, mainHeader, mainHeaderRow, inputFile.getPath() + ".xml");
			XLSUtil.saveXLS(wb, inputFile);
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
	private static boolean updateGeoDataInExcel(Workbook wb, int mainSheet, int mainRow,
			Map<String, Integer> mainHeader, int mainHeaderRow, HashMap<String, Integer> calculatedLoc) {
		boolean hasUpdated = false;
		
		//Find the secondary sheet
		int secSheet = XLSUtil.getOrCreateSheet(wb, "Secondary", 1);
		if (secSheet == 0)
			secSheet = XLSUtil.getOrCreateSheet(wb, "Secondary 2", 1);

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
		
		boolean failedQuery = false;
		//Scan all locations query them and output the results
		while (!XLSUtil.isEndRow(wb, mainSheet, mainRow) && failedQuery == false) {
			String location = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LOCATION));
			String isProcessed = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_IS_PROCESSED));
			// Check if row needs to be written
			if ((calculatedLoc.get(location) != null) && (Boolean.valueOf(isProcessed) == false))
				setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_IS_PROCESSED, "Already calculated (" + Integer.toString(calculatedLoc.get(location) + 1) + ")");
			else if ((location.isEmpty() == false) && (Boolean.valueOf(isProcessed) == false)) {
				List<GeocoderResult> queryRes = geoQuerier.makeQuery(location);
				if (queryRes == null)
				{
					System.out.println("Quit due to failed query.");
					failedQuery = true;
				}
				else {
					// Write the query result
					boolean hasResult = !queryRes.isEmpty();
					setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_IS_PROCESSED, true);
					setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_HAS_RESULT,
							Boolean.toString(hasResult));
					if (queryRes.size() > 1)
					{
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_SEC_RESULT_START, secRow + 1);
						setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_SEC_RESULT_END, secRow + queryRes.size());
					}
					writeGeoResultToRow(wb, mainSheet, mainRow, mainHeader, hasResult ? queryRes.get(0) : null, true);
					

					// Write the non-main result
					for (int i = 1; i < queryRes.size(); ++i) {
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

	private static HashMap<String, Integer> getPreCalculatedLocations(Workbook wb, int mainSheet, int mainRow,
			Map<String, Integer> mainHeader) {
		HashMap<String, Integer> precalced = new HashMap<String, Integer>();
		
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
					.getLat().toString());
			setCellByHeader(wb, sheetNum, readRow, workHeader, COLUMN_LONGITUDE, mainResult.getGeometry().getLocation()
					.getLng().toString());
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
			String columnName, boolean value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}
	private static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, int value) {
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
				isValidHeader &= tmpColumns.get(neededColumn) != null;
			if (isValidHeader)
				headerColumns.putAll(tmpColumns);
		} while ((isValidHeader == false) && (XLSUtil.isEndRow(wb, sheetNum, rowNum) == false));
		
		if (isValidHeader == true)
			return rowNum;
		else return -1;
	}
	

	private static boolean exportToXML(final Workbook wb, final int mainSheet, final int mainStartRow, final Map<String, Integer> mainHeader, int mainHeaderRow, String fileName) {
		//Find the secondary sheet
		int xmlSheet = XLSUtil.getOrCreateSheet(wb, "XML Settings", 2);
		
		HashMap<String,Integer> xmlHeader = new HashMap<>();
		int xmlHeaderRow = getHeaderRow(wb, xmlSheet, xmlHeader, COLUMN_XML_LEVEL, COLUMN_XML_NAME, COLUMN_XML_FROM);
		if (xmlHeaderRow == -1)
		{
			System.out.println("No XML header with rows \"" + COLUMN_XML_LEVEL + "\", \"" + COLUMN_XML_NAME +
					"\" and \"" + COLUMN_XML_FROM + "\" found in XML Settings sheet.");
		}
		else
		{
			int xmlRow = xmlHeaderRow + 1;
			
			//Scan the list of xml levels
			ArrayList<String> names = new ArrayList<String>();
			ArrayList<String> froms = new ArrayList<String>();
			while (!XLSUtil.isEndRow(wb, xmlSheet, xmlRow)) {
				
				String name = XLSUtil.getCellString(wb, xmlSheet, xmlRow, xmlHeader.get(COLUMN_XML_NAME));
				String from = XLSUtil.getCellString(wb, xmlSheet, xmlRow, xmlHeader.get(COLUMN_XML_FROM));
				if ((!name.isEmpty()) || (!from.isEmpty()))
				{
					names.add(name);
					froms.add(from);
				}
				++xmlRow;
			}
			
			//Gather all rows with info in main xml
			int mainRow = mainStartRow;
			ArrayList<Integer> rowsWithInfo = new ArrayList<>();
			while (!XLSUtil.isEndRow(wb, mainSheet, mainRow)) {
				String types = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_TYPE));
				int wordIndex = containsWordFromListIndex(types, froms);
				if (wordIndex != -1)
					rowsWithInfo.add(new Integer(mainRow));
				++mainRow;
			}
			
			final ArrayList<String> finalFroms = froms;
			//Order the rows
			Collections.sort(rowsWithInfo, new Comparator<Integer>() {

				@Override
				public int compare(Integer o1, Integer o2) {
					int comp = 0;
					for(String from : finalFroms)
					{
						String val1 = XLSUtil.getCellString(wb, mainSheet, o1, mainHeader.get(COLUMN_GOOGLE_TYPE_PREFIX + from));
						String val2 = XLSUtil.getCellString(wb, mainSheet, o2, mainHeader.get(COLUMN_GOOGLE_TYPE_PREFIX + from));
						comp = val1.compareTo(val2);
						if (comp != 0)
							break;
					}
					return comp;
				}});
			
			//Write the XML to a file			
			try {
				File xmlFilename = new File(fileName);

				DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

				// root elements
				Document doc = docBuilder.newDocument();
				Element rootElement = doc.createElement("MultiLayerDictionary");
				doc.appendChild(rootElement);

				Element levels = doc.createElement("Levels");
				rootElement.appendChild(levels);
				for(String name : names)
				{
					Element level = doc.createElement("Level");
					level.setAttribute("name", name);
					rootElement.setAttribute("type", "Location");
					levels.appendChild(level);
				}
				
				Element entries = doc.createElement("Entries");
				rootElement.appendChild(entries);
				
				ArrayList<String> prevValues = new ArrayList<String>(Collections.nCopies(froms.size(), (String)null));
				List<Element> elements = new ArrayList<Element>(Collections.nCopies(froms.size(), (Element)null));
				for(Integer row : rowsWithInfo)
				{
					//Check which values have changed
					boolean hasActualValue = false;
					for(int i = froms.size() - 1 ; i >= 0 ; --i)
					{
						String value = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_GOOGLE_TYPE_PREFIX + froms.get(i)));
						if (value.equals(prevValues.get(i)) == false)
						{
							hasActualValue |= !value.isEmpty();
							if (hasActualValue)
								prevValues.set(i, value);
							else prevValues.set(i, null);
							elements = elements.subList(0, i);
						}
					}
					
					//Add the new values to the 
					
					Element lastEntryToBeAdded = null;
					for(int i = elements.size() ; i <  froms.size() ; ++i)
					{
						if (prevValues.get(i) != null)
						{
							Element entry = doc.createElement("Entry");
							entry.setAttribute("name", prevValues.get(i));
							
							Element parent = i == 0 ? entries : elements.get(i - 1);
							parent.appendChild(entry);
							elements.add(entry);
							lastEntryToBeAdded = entry;
						}
					}
					
					if (lastEntryToBeAdded != null)
					{
						String altNames = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_ALTERNATIVE_NAME));
						if (altNames.isEmpty() == false)
							lastEntryToBeAdded.setAttribute("standAloneNames", altNames);
						String groupStr = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_GROUP));
						if (groupStr.isEmpty() == false)
							lastEntryToBeAdded.setAttribute("group", groupStr);
					}
					
					/*String longitudeStr = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_LONGITUDE));
					if (longitudeStr.isEmpty() == false)
						entry.setAttribute("longitude", longitudeStr);
					String latitudeStr = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_LONGITUDE));
					if (latitudeStr.isEmpty() == false)
						entry.setAttribute("latitude", latitudeStr);*/
				}
				
				// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer = transformerFactory.newTransformer();
				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new PrintStream(xmlFilename));
				transformer.setOutputProperty(OutputKeys.INDENT, "yes");
				transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
				transformer.transform(source, result);
				
				//Reupdate the row numbering
				if (mainHeader.get(COLUMN_ORIG_ROW) != null)
				{
					mainRow = mainStartRow;
					while (XLSUtil.isEndRow(wb, mainSheet, mainRow) == false)
					{
						String location = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LOCATION));
						String rowNum = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_ORIG_ROW));
						if ((location.isEmpty() == false) && (rowNum.isEmpty() == false))
							setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_ORIG_ROW, "");
						++mainRow;
					}
				}
				
				int rowCount = 0;
				for(Integer row : rowsWithInfo)
				{
					setCellByHeader(wb, mainSheet, row, mainHeader, COLUMN_ORIG_ROW, ++rowCount);
				}
				
				XLSUtil.updateHeader(wb, mainSheet, mainHeaderRow, mainHeader);
				
				
				
			} catch (Exception e) {
				e.printStackTrace();
				return false;
			}
		}

		return true;
	}

	/**
	 * @param list
	 * @param string
	 * @return
	 */
	private static int containsWordFromListIndex(String string, ArrayList<String> list) {
		for(int i = 0 ; i < list.size() ; ++i)
		{
			if (string.matches("(?i:.*\\b" + list.get(i) + "\\b.*)"))
				return i;
		}
		return -1;
	}

}
