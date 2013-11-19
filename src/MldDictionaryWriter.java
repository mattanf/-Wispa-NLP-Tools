
import java.io.File;
import java.io.PrintStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.google.code.geocoder.model.GeocoderAddressComponent;
import com.google.code.geocoder.model.GeocoderResult;
import com.google.code.geocoder.model.LatLng;

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

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

public class MldDictionaryWriter {

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
	final static String COLUMN_FLAGS = "Flags";
	final static String COLUMN_COUNTRY_REGION = "CountryRegion";
	
	final static String COLUMN_XML_LEVEL = "Level";
	final static String COLUMN_XML_NAME = "Name";
	final static String COLUMN_XML_FROM = "From";
	
	final static String PROPERTY_FILENAME = "FileName";
	final static String PROPERTY_DATA_TYPE = "DataType";
	final static String PROPERTY_DEFAULT_SEARCH_FLAGS = "DefaultSearchFlags";
	
	final static String SHEET_REGION_BY_CLOSNESS = "RegionByClosness";
	final static String PROPERTY_COLUMN_TO_ADD = "ColumnToAdd";
	final static String PROPERTY_COLUMN_TO_COMPARE = "ColumnToCompare";
	final static String COLUMN_CLOSE_BY_REGION = "CloseByRegion";
	
	
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
		
		if (mainHeaderRow == -1) {
			System.out.println("No header with rows \"" + COLUMN_LOCATION + "\" and \"" + COLUMN_IS_PROCESSED +
					"\" found in main sheet.");
		} else {
			
			MldEntryType dataType = MldEntryType.toEnum(getPreHeaderProperty(wb, mainSheet, mainHeaderRow, PROPERTY_DATA_TYPE));
			if (dataType == null)
				dataType = MldEntryType.General;
			addColumnToHeader(mainHeader, COLUMN_ORIG_ROW);
			addColumnToHeader(mainHeader, COLUMN_TYPE);
			addColumnToHeader(mainHeader, COLUMN_ALTERNATIVE_NAME);
			addColumnToHeader(mainHeader, COLUMN_FLAGS);
			
			HashMap<String, Integer> calculatedLoc = getPreCalculatedLocations(wb, mainSheet, mainHeaderRow, mainHeader);
			
			boolean hasUpdated = false;
			if (dataType == MldEntryType.Location)
			{
				addColumnToHeader(mainHeader, COLUMN_LATITUDE);
				addColumnToHeader(mainHeader, COLUMN_LONGITUDE);
				hasUpdated = updateGeoDataInExcel(wb, mainSheet, mainHeader, mainHeaderRow, calculatedLoc);
				hasUpdated |= updateClosestMajorCity(wb, mainSheet, mainHeader, mainHeaderRow);
			}
			
			hasUpdated |= exportToXML(wb, dataType, mainSheet, mainHeader, mainHeaderRow, new File(inputFile.getPath() + ".xml"));
			if (hasUpdated)
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
	private static boolean updateGeoDataInExcel(Workbook wb, int mainSheet, 
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
		int mainRow = mainHeaderRow + 1;
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

	private static HashMap<String, Integer> getPreCalculatedLocations(Workbook wb, int mainSheet, int mainHeaderRow,
			Map<String, Integer> mainHeader) {
		HashMap<String, Integer> precalced = new HashMap<String, Integer>();
		
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

	private static boolean updateClosestMajorCity(Workbook wb, int mainSheet, 
			Map<String, Integer> mainHeader, int mainHeaderRow) {
		boolean isChanged = false;
		int citiesSheet = XLSUtil.getSheetNumber(wb, SHEET_REGION_BY_CLOSNESS);
		if (citiesSheet == -1)
			System.out.println("Sheet " + SHEET_REGION_BY_CLOSNESS + " not found.");
		else
		{
			Map<String, Integer> citiesHeader = new HashMap<>();
			Map<String, Map<String, LatLng>> mapPositionToLocation = new HashMap<>();
			int cityRow = getHeaderRow(wb, citiesSheet, citiesHeader, COLUMN_LATITUDE, COLUMN_LONGITUDE);
			String columnToAdd = getPreHeaderProperty(wb, citiesSheet, cityRow, PROPERTY_COLUMN_TO_ADD);
			String columnToCompare = getPreHeaderProperty(wb, citiesSheet, cityRow, PROPERTY_COLUMN_TO_COMPARE);
			if ((columnToCompare != null) && (columnToCompare.isEmpty()))
				columnToCompare = null;
			if ((columnToAdd != null) && (columnToAdd.isEmpty()))
				columnToAdd = null;
			
			if (cityRow == -1)
				System.out.println("Header of sheet " + SHEET_REGION_BY_CLOSNESS + " not found. Cannot perform closest region update.");
			else if (columnToAdd == null)
				System.out.println("ColumnToAdd not found in sheet " + SHEET_REGION_BY_CLOSNESS + ". Cannot perform closest region update.");
			else if (citiesHeader.get(columnToAdd) == null)
				System.out.println("ColumnToAdd " + columnToAdd + " not found in header of sheet " + SHEET_REGION_BY_CLOSNESS + ". Cannot perform closest region update.");
			else if ((columnToCompare != null) && (citiesHeader.get(columnToCompare) == null))
				System.out.println("ColumnToCompare " + columnToCompare + " not found in header of sheet " + SHEET_REGION_BY_CLOSNESS + ". Cannot perform closest region update.");
			else if ((columnToCompare != null) && (mainHeader.get(columnToCompare) == null))
				System.out.println("ColumnToCompare " + columnToCompare + " not found in header of main sheet. Cannot perform closest region update.");
			else if (mainHeader.get(COLUMN_CLOSE_BY_REGION) == null)
				System.out.println("Column " + COLUMN_CLOSE_BY_REGION + " not found in header of main sheet. Cannot perform closest region update.");
			else
			{
				++cityRow;
				while (!XLSUtil.isEndRow(wb, citiesSheet, cityRow))
				{
					try{
						
						String name = XLSUtil.getCellString(wb, citiesSheet, cityRow, citiesHeader.get(columnToAdd));
						String compareValue = null;
						if (columnToCompare != null)
							compareValue = XLSUtil.getCellString(wb, citiesSheet, cityRow, citiesHeader.get(columnToCompare));
						String latitude = XLSUtil.getCellString(wb, citiesSheet, cityRow, citiesHeader.get(COLUMN_LATITUDE));
						String longitude = XLSUtil.getCellString(wb, citiesSheet, cityRow, citiesHeader.get(COLUMN_LONGITUDE));
						if (name.isEmpty() == false)
						{
							if (!mapPositionToLocation.containsKey(compareValue))
								mapPositionToLocation.put(compareValue, new HashMap<String, LatLng>());
							mapPositionToLocation.get(compareValue).put(name, new LatLng(latitude, longitude));
						}
							
					}
					catch (Throwable e)
					{
						System.out.println("Unable to parse Major cities line " + Integer.toString(cityRow + 1));
					}
					++cityRow;
				}
				
				if (mapPositionToLocation.isEmpty() == false)
				{
					Integer columnIndex = addColumnToHeader(mainHeader, COLUMN_CLOSE_BY_REGION);
					
					int mainRow = mainHeaderRow + 1;
					while (!XLSUtil.isEndRow(wb, mainSheet, mainRow))
					{
						try {
							String latitude = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LATITUDE));
							String longitude = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LONGITUDE));
							String compareValue = null;
							if (!latitude.isEmpty() && !longitude.isEmpty())
							{
								if (columnToCompare != null)
									compareValue = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(columnToCompare));
								Map<String, LatLng> mapInCompared = mapPositionToLocation.get(compareValue);
								String newValue = null;
								if (mapInCompared != null)
								{
									LatLng location = new LatLng(latitude, longitude);
									newValue = getClosestLocationName(location, mapInCompared);
								}
								String prevValue = XLSUtil.getCellString(wb, mainSheet, mainRow, columnIndex);
								if (prevValue != newValue)
								{
									XLSUtil.setCellString(wb, mainSheet, mainRow, columnIndex, newValue);
									isChanged = true;
								}
							}
						}
						catch(Throwable e)
						{
							System.out.println("Unable to save closest major city in line " + Integer.toString(mainRow + 1));
						}
						++mainRow;
					}
				}
				
			}
		}
		return isChanged;
	}

	private static String getClosestLocationName(LatLng location, Map<String, LatLng> mapPositionToLocation) {
		double closestSqrDistance = Double.MAX_VALUE;
		String closestLocationName = null;
		for(Entry<String, LatLng> entry : mapPositionToLocation.entrySet())
		{
			double latDiff = entry.getValue().getLat().doubleValue() - location.getLat().doubleValue();
			double lngDiff = entry.getValue().getLng().doubleValue() - location.getLng().doubleValue();
			double entSqrDist = latDiff * latDiff + lngDiff * lngDiff;
			if (entSqrDist < closestSqrDistance)
			{
				closestSqrDistance = entSqrDist;
				closestLocationName = entry.getKey();
			}
		}
		return closestLocationName;
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
	

	private static boolean exportToXML(final Workbook wb, MldEntryType dataType, final int mainSheet, 
			final Map<String, Integer> mainHeader, int mainHeaderRow, File fileName) {
		//Find the secondary sheet
		int xmlSheet = XLSUtil.getOrCreateSheet(wb, "XML Settings", 2);
		boolean hasDataChanged = false;
		
		HashMap<String,Integer> xmlHeader = new HashMap<>();
		int xmlHeaderRow = getHeaderRow(wb, xmlSheet, xmlHeader, COLUMN_XML_LEVEL, COLUMN_XML_NAME, COLUMN_XML_FROM);
		if (xmlHeaderRow == -1)
		{
			System.out.println("No XML header with rows \"" + COLUMN_XML_LEVEL + "\", \"" + COLUMN_XML_NAME +
					"\" and \"" + COLUMN_XML_FROM + "\" found in XML Settings sheet.");
		}
		else
		{
			String defaultSearchFlags = getPreHeaderProperty(wb, mainSheet, mainHeaderRow, PROPERTY_DEFAULT_SEARCH_FLAGS);
			String newFileNameTmp = getPreHeaderProperty(wb, xmlSheet, xmlHeaderRow, PROPERTY_FILENAME);
			if (newFileNameTmp != null)
			{
				File newFileName = new File(newFileNameTmp);
				if ((!newFileName.isAbsolute()) && (fileName.getParent() != null))
					fileName = new File(fileName.getParent() + "\\" + newFileName.getPath());
			}
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
			int mainRow = mainHeaderRow + 1;
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
						String val1 = XLSUtil.getCellString(wb, mainSheet, o1, mainHeader.get(from));
						String val2 = XLSUtil.getCellString(wb, mainSheet, o2, mainHeader.get(from));
						comp = val1.compareTo(val2);
						if (comp != 0)
							break;
					}
					return comp;
				}});
			
			//Write the XML to a file			
			try {
				File xmlFilename = new File(fileName.getPath());

				DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
				DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

				// root elements
				Document doc = docBuilder.newDocument();
				Element rootElement = doc.createElement("MultiLayerDictionary");
				doc.appendChild(rootElement);
				rootElement.setAttribute("type", dataType.name());
				rootElement.setAttribute("defaultSearchFlags", defaultSearchFlags);
				

				Element levels = doc.createElement("Levels");
				rootElement.appendChild(levels);
				for(String name : names)
				{
					Element level = doc.createElement("Level");
					level.setAttribute("name", name);
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
						String value = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(froms.get(i)));
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
						String tmpVal;
						tmpVal = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_ALTERNATIVE_NAME));
						if (tmpVal.isEmpty() == false)
							lastEntryToBeAdded.setAttribute("standAloneNames", tmpVal);
						tmpVal = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_FLAGS));
						if (tmpVal.isEmpty() == false)
							lastEntryToBeAdded.setAttribute("flags", tmpVal);
						if (dataType == MldEntryType.Location)
						{
							try {
							tmpVal = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_LONGITUDE));
							if (tmpVal.isEmpty() == false)
								lastEntryToBeAdded.setAttribute("longitude", String.format("%.4f", Double.valueOf(tmpVal)));
							tmpVal = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_LATITUDE));
							if (tmpVal.isEmpty() == false)
								lastEntryToBeAdded.setAttribute("latitude", String.format("%.4f", Double.valueOf(tmpVal)));
							}
							catch(Throwable e)
							{
								System.out.println("Unable to write long lat");
							}
						}
							
						/*if (dataType == MldDataType.Location)
						{
							String cmlValue = XLSUtil.getCellString(wb, mainSheet, row, mainHeader.get(COLUMN_COUNTRY_REGION));
							if (cmlValue.isEmpty() == false)
								lastEntryToBeAdded.setAttribute("majorCity", cmlValue);
						}*/
					}
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
					mainRow = mainHeaderRow + 1;
					while (XLSUtil.isEndRow(wb, mainSheet, mainRow) == false)
					{
						int newOrderNum = rowsWithInfo.indexOf(new Integer(mainRow)) + 1;
						String orderNum = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_ORIG_ROW));
						if (!orderNum.equals(newOrderNum == 0 ? "" : Integer.toString(newOrderNum)))
						{
							hasDataChanged = true;
							setCellByHeader(wb, mainSheet, mainRow, mainHeader, COLUMN_ORIG_ROW, newOrderNum == -1 ? (Integer)null : new Integer(newOrderNum));
						}
						++mainRow;
					}
				}
								
				XLSUtil.updateHeader(wb, mainSheet, mainHeaderRow, mainHeader);
			} catch (Exception e) {
				e.printStackTrace();
				return false;
			}
		}

		return hasDataChanged;
	}

	private static String getPreHeaderProperty(Workbook wb, int sheetNum, int headerRow, String propertyName) {
		int curRow = 0;
		while ((curRow < headerRow) && (!XLSUtil.isEndRow(wb, sheetNum, curRow)))
		{
			String curPropName = XLSUtil.getCellString(wb, sheetNum, curRow, 0);
			if (propertyName.equalsIgnoreCase(curPropName))
			{
				return XLSUtil.getCellString(wb, sheetNum, curRow, 1);
			}
			++curRow;
		}
		return null;
	}

	/**
	 * @param list
	 * @param string
	 * @return
	 */
	private static int containsWordFromListIndex(String string, ArrayList<String> list) {
		for(int i = 0 ; i < list.size() ; ++i)
		{
			String listItem = list.get(i);
			if (listItem.startsWith(COLUMN_GOOGLE_TYPE_PREFIX))
			{	
				listItem = listItem.substring(COLUMN_GOOGLE_TYPE_PREFIX.length());
				if (string.matches("(?i:.*\\b" + listItem + "\\b.*)"))
					return i;
			}
		}
		return -1;
	}

}
