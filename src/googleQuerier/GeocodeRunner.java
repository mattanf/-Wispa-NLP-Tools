package googleQuerier;

import googleQuerier.GeocodeDataManager.QueryData;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.code.geocoder.model.GeocoderResult;
import com.google.code.geocoder.model.LatLng;
import com.pairapp.engine.parser.location.LocationLanguageData;
import com.pairapp.engine.parser.location.LocationLanguageData.MisspellCorrectionType;
import com.pairapp.engine.parser.location.NamedLocation;
import com.pairapp.engine.parser.location.NamedLocation.LocationType;

import utility.XLSUtil;

public class GeocodeRunner {

	private static final String GEOCODE_OUTPUT_DIR = "GeocodeData";
	private static final String DEFAULT_COUNTRY_NAME = "ישראל";
	private static final int MIN_STREET_COUNT_FOR_INCLUDE = 15;
	private static ArrayList<LocationEntry> cityByDistance = null;
	private static XSSFWorkbook streetWB;
	private static String workingDir;
	private static HashMap<String, ArrayList<NamedLocation>> manualLocations = null;
	
	/**
	 * This program recives an xlsx file that contains a list of location. It queries google's geocode web service for
	 * information, saves the information to the disk, than later updates the xlsx file with the information,
	 * 
	 * @param args
	 * @throws IOException
	 * @throws ClassNotFoundException
	 */
	public static void main(String[] args) throws ClassNotFoundException, IOException {

		String fileToMergeFrom = null;
		boolean gatherDataFromGoogle = false;
		if ((args.length > 0) && (args[0].equals("-scan")))
			gatherDataFromGoogle = true;
		if ((args.length > 1) && (args[0].equals("-merge")))
			fileToMergeFrom = args[1];

		manualLocations = new HashMap<String, ArrayList<NamedLocation>>();
		workingDir = new File(".").getCanonicalPath();
		if (new File(workingDir, "rehovot.xlsx").exists())
			streetWB = XLSUtil.openXLS("rehovot.xlsx");
		else {
			workingDir = "C:\\Public\\Train Data\\";
			if (new File(workingDir, "rehovot.xlsx").exists())
				streetWB = XLSUtil.openXLS("C:\\Public\\Train Data\\rehovot.xlsx");
		}

		
		if (streetWB == null) {
			System.out.println("Error: Unable to find rehovot.xlsx source file.");
		} else {
			File outDir = new File(workingDir, GEOCODE_OUTPUT_DIR);
			if (outDir.exists() == false)
				outDir.mkdir();
			workingDir = outDir.getCanonicalPath();
			GeocodeDataManagerCollection.initialize(workingDir);
			
			HashMap<String, String> cityToCode = new HashMap<>();
			scanCityInfo(cityToCode);
			if (gatherDataFromGoogle) {
				scanStreetInfo();
			}
			mergeFromAdditionalFile(fileToMergeFrom);
			outputGatheredData(cityToCode,fileToMergeFrom);
			
		}
		
		GeocodeDataManagerCollection.deinitialize(workingDir);
		System.out.println("Work done.");
	}

	private static void mergeFromAdditionalFile(String fileToMergeFrom) {
		XSSFWorkbook mergeWB = XLSUtil.openXLS(fileToMergeFrom);
		int sheetNum = XLSUtil.getSheetNumber(mergeWB, "Locations");
		if (("Latitude".equals(XLSUtil.getCellString(mergeWB, sheetNum, 0, 2))) &&
				("Longitude".equals(XLSUtil.getCellString(mergeWB, sheetNum, 0, 3))))
		{
			boolean hasChanges = false;
			int readRow = 1;
			while (!XLSUtil.isEndRow(mergeWB, sheetNum, readRow))
			{
				String locationName = XLSUtil.getCellString(mergeWB, sheetNum, readRow, 0);
				String status = XLSUtil.getCellString(mergeWB, sheetNum, readRow, 4);
				if ((locationName.isEmpty() == false) && (status.equals("") || status.equals("Fail")))
				{
					GeocoderResult res = GeocodeDataManagerCollection.getAddressGeodata(locationName, true, false);
					hasChanges = true;
					if (res == null)
					{
						XLSUtil.setCellString(mergeWB, sheetNum, readRow, 4, "Fail");
					}
					else
					{
						XLSUtil.setCellString(mergeWB, sheetNum, readRow, 2, res.getGeometry().getLocation().getLat().toString());
						XLSUtil.setCellString(mergeWB, sheetNum, readRow, 3, res.getGeometry().getLocation().getLng().toString());
						XLSUtil.setCellString(mergeWB, sheetNum, readRow, 4, "Pass");
					}
				}
				if ((locationName.isEmpty() == false) && (status.equals("Manual")))
				{
					int delim = locationName.indexOf(", ");
					LocationType locType = LocationType.toEnum(XLSUtil.getCellString(mergeWB, sheetNum, readRow, 1));
					Double lat = new Double(Double.parseDouble(XLSUtil.getCellString(mergeWB, sheetNum, readRow, 2)));
					Double lng = new Double(Double.parseDouble(XLSUtil.getCellString(mergeWB, sheetNum, readRow, 3)));
					if ((locType != null) && (lat != null) && (lng != null) && (delim != -1))
					{
						String parentName = locationName.substring(delim + 2);
							locationName =  locationName.substring(0, delim);
						if (manualLocations.get(parentName) == null)
							manualLocations.put(parentName, new ArrayList<NamedLocation>());
						manualLocations.get(parentName).add(new NamedLocation(locType, null, null, locationName, null, lat, lng));
					}
				}
									
				++readRow;
			}
			
			if (hasChanges)
				XLSUtil.saveXLS(mergeWB,fileToMergeFrom);
		}	
		else System.out.println("Error - no proper merge file found");
	}

	private static void outputGatheredData(HashMap<String, String> cityToCode, String languageExcel) {
		LocationLanguageData langData = readLanuguageDataFromFile(languageExcel);
		GeocodeDataToXml serializer = new GeocodeDataToXml(langData);
		
		ArrayList<NamedLocation> citiesToWrite = new ArrayList<NamedLocation>();
		Iterator<GeocoderResult> cityIt = GeocodeDataManagerCollection.getCountryData().getUniqueGeoResults();
		while (cityIt.hasNext()) {
			GeocoderResult cityRes = cityIt.next();
			NamedLocation parentEntity = geoResultToNamedEntity(cityRes,GeocodeDataManagerCollection.DEFAULT_COUNTRY_LOCATION);
			if (parentEntity != null)
			{
				citiesToWrite.add(parentEntity);
			
				GeocodeDataManager man = GeocodeDataManagerCollection.getGeocodeManager(cityRes.getFormattedAddress());
				if ((man != null) && (man.getUniqueResultCount() > MIN_STREET_COUNT_FOR_INCLUDE)) {
					
					String parentName = man.getBaseLocationName();
					String locationCode = cityToCode.get(parentName);
					if (locationCode == null)
					{
						locationCode = generateCodeIfNeeded(parentName);
						cityToCode.put(parentName, locationCode);
					}
					if (locationCode == null)
					{
						locationCode = parentName;
					}
					locationCode = GeocodeDataManagerCollection.DEFAULT_COUNTRY_CODE + "-" + locationCode;
					ArrayList<NamedLocation> subCityLocations = geoResultsToNamedLocationArray(man.getUniqueGeoResults(), parentEntity);
					subCityLocations = geoResultsToNamedLocationArray(man.getUniqueGeoResults(), parentEntity);
					subCityLocations = mergeAutomaticAndManualLocations(parentName, subCityLocations);
					serializer.writeXMLtoFile(parentName, locationCode, subCityLocations.iterator(), workingDir,
							LocationType.City, null);
					
					
				}
			}
			
		}		
		
		String countryName = GeocodeDataManagerCollection.getCountryData().getBaseLocationName();
		citiesToWrite = mergeAutomaticAndManualLocations(countryName, citiesToWrite);
		serializer.writeXMLtoFile(countryName, GeocodeDataManagerCollection.DEFAULT_COUNTRY_CODE,
				citiesToWrite.iterator(), workingDir,
				LocationType.Country, cityToCode);
	}
	
		
	private static ArrayList<NamedLocation> mergeAutomaticAndManualLocations(String parentName,
			ArrayList<NamedLocation> citiesToWrite) {
		
		ArrayList<NamedLocation> arrayList = manualLocations.get(parentName);
		if (arrayList != null)
		{
			citiesToWrite.addAll(arrayList);
		}
		return citiesToWrite;
	}

	private static ArrayList<NamedLocation> geoResultsToNamedLocationArray(Iterator<GeocoderResult> uniqueGeoResults,
			NamedLocation parentLocation) {
		
		ArrayList<NamedLocation> retArray = new ArrayList<>();
		while (uniqueGeoResults.hasNext())
		{
			GeocoderResult res = uniqueGeoResults.next();
			NamedLocation retEntity = geoResultToNamedEntity(res, parentLocation);
			if (retEntity != null)
				retArray.add(retEntity);
		}
		return retArray;
	}

	/**
	 * @param res
	 * @param parentLocation
	 * @return
	 */
	private static NamedLocation geoResultToNamedEntity(GeocoderResult res, NamedLocation parentLocation) {
		NamedLocation retEntity = null; 
		LocationType curLocType = getLocationType(res);
		String addressName = res.getAddressComponents().get(0).getLongName();
		if  (curLocType != null)
		{	
			Double lat = null;
			Double lng = null;
			if ((res.getGeometry() != null) && (res.getGeometry().getLocation() != null) &&
					(res.getGeometry().getLocation().getLat() != null)) {
				lat = new Double(res.getGeometry().getLocation().getLat().doubleValue());
				lng = new Double(res.getGeometry().getLocation().getLng().doubleValue());
			}
			retEntity = new NamedLocation(curLocType,null,parentLocation,addressName, null,lat, lng);
		}
		return retEntity;
	}
	

	private static LocationType getLocationType(GeocoderResult res) {
		String curType = res.getTypes().size() > 0 ? res.getTypes().get(0) : null;
		if (curType != null) {
			if (curType.equals("locality"))
				return LocationType.City;
			else if (curType.equals("route"))
				return LocationType.Street;
			else if (curType.equals("neighborhood"))
				return LocationType.CityRegion;
			else if (curType.equals("park"))
				return LocationType.CityRegion;
			else if (curType.equals("point_of_interest"))
				return LocationType.CityRegion;
			else if (curType.equals("establishment"))
				return LocationType.CityRegion;
			else
				return null;
		}
		return null;
	}

	private static String generateCodeIfNeeded(String cityName) {
		boolean isAscii = cityName.matches("\\p{ASCII}*");
		if (!isAscii)
		{
			GeocoderResult res = GeocodeDataManagerCollection.getAddressGeodata(cityName, false, false);
			int lat = ((int)(res.getGeometry().getLocation().getLat().doubleValue() * 100) + 36000) % 36000;
			int lng = ((int)(res.getGeometry().getLocation().getLng().doubleValue() * 100) + 36000) % 36000;
			return String.format("%1$05d%2$05d", lat, lng);
		}
		return null;		
	}

	private static LocationLanguageData readLanuguageDataFromFile(String languageExcel) {
		//מיותר	תחילית	שם אב	ניתן לנירמול	שם 1	שם 2	שם 3	שם 4
		final int LANG_COL_TYPE = 0;
		final int LANG_COL_REDUNDENT = 1;
		final int LANG_COL_PARENT = 2;
		final int LANG_COL_ALT_NAME_BEGIN = 3;
		
		final int LANG_COL_EXCLUDED = 0;
		final int LANG_COL_ISREGEX = 1;
		
		XSSFWorkbook langWB = XLSUtil.openXLS(languageExcel);
		int misspellsSheet = XLSUtil.getSheetNumber(langWB, "Misspells");
		int excludedSheet = XLSUtil.getSheetNumber(langWB, "Excluded Names");
		if ((misspellsSheet != -1) || (excludedSheet != -1))
		{
			LocationLanguageData retLangData = new LocationLanguageData("he");
			int readRow = 1;
			while (!XLSUtil.isEndRow(langWB, misspellsSheet, readRow))
			{
				boolean isRedundent = Boolean.valueOf(XLSUtil.getCellString(langWB, misspellsSheet, readRow, LANG_COL_REDUNDENT));
				String typeStr = XLSUtil.getCellString(langWB, misspellsSheet, readRow, LANG_COL_TYPE);
				
				MisspellCorrectionType type = MisspellCorrectionType.toEnum(typeStr);
				if (type == null)
					System.out.println("Unknown type in misspel sheet on line " + Integer.toString(readRow));
				
				String parent = XLSUtil.getCellString(langWB, misspellsSheet, readRow, LANG_COL_PARENT);
				if (parent.isEmpty())
					parent = null;
				
				ArrayList<String> arr = new ArrayList<>();
				for(int i = LANG_COL_ALT_NAME_BEGIN ; i < 100 ; ++i)
				{
					String altName = XLSUtil.getCellString(langWB, misspellsSheet, readRow, i);
					if (!altName.isEmpty())
						arr.add(altName);
					else break;
				}
				
				if ((arr.size() > 0) && (type != null))
				{
					retLangData.addMisspellCorrection(type, arr.toArray(new String[arr.size()]), isRedundent, parent); 
				}

				++readRow;
			}
			
			readRow = 1;
			while (!XLSUtil.isEndRow(langWB, excludedSheet, readRow))
			{
				String value = XLSUtil.getCellString(langWB, excludedSheet, readRow, LANG_COL_EXCLUDED);
				boolean isRegEx = Boolean.valueOf(XLSUtil.getCellString(langWB, excludedSheet, readRow, LANG_COL_ISREGEX));
				if (!value.isEmpty())
					retLangData.addExcludedString(value, isRegEx);
				++readRow;
			}
			return retLangData;
		}		
		else System.out.println("Error - no proper merge file found");
		
		return null;
	}	
	

	private static void scanCityInfo(HashMap<String, String> cityToCode) {

		cityByDistance = new ArrayList<LocationEntry>();
		int readRow = 1;
		boolean userBreak = false;
		LocationEntry prevLoc = null;
		int countCityRepeat = 0;
		while (!XLSUtil.isEndRow(streetWB, readRow) && !userBreak) {
			LocationEntry loc = new LocationEntry(streetWB, readRow);
			++readRow;
			loc = loc.generateCityLocation();
			if (loc.equals(prevLoc))
			{	
				++countCityRepeat;
				if ((countCityRepeat >= MIN_STREET_COUNT_FOR_INCLUDE) && 
						(cityByDistance.contains(loc) == false)) {
					GeocoderResult res = GeocodeDataManagerCollection.getAddressGeodata(getEntryFormattedAdress(loc, null), true, false);
					if ((res != null) && (res.getTypes() != null) && (res.getTypes().size() > 0) &&
							(res.getTypes().get(0).equals("locality")))
					{
						cityToCode.put(res.getFormattedAddress(), loc.getCityId());
						cityByDistance.add(loc);
					}
				}
			}
			else countCityRepeat = 1;
			prevLoc = loc;
		}

		QueryData scanFromRes1 = GeocodeDataManagerCollection.getCountryData().getOrCreateQueryData("תל אביב יפו, ישראל");
		QueryData scanFromRes2 = GeocodeDataManagerCollection.getCountryData().getOrCreateQueryData("ירושלים, ישראל");
		QueryData scanFromRes3 = GeocodeDataManagerCollection.getCountryData().getOrCreateQueryData("באר שבע, ישראל");
		QueryData scanFromRes4 = GeocodeDataManagerCollection.getCountryData().getOrCreateQueryData("חיפה, ישראל");
		LatLng telHai = new LatLng("33.2340070", "35.5795340");
		//scanFromRes3 = scanFromRes2 = scanFromRes1;
		if ((scanFromRes1 != null) && (scanFromRes1.getDescription().isRecordExists())) {
			final LatLng relLatLng1 = scanFromRes1.getResults()[0].getGeometry().getLocation();
			final LatLng relLatLng2 = scanFromRes2.getResults()[0].getGeometry().getLocation();
			final LatLng relLatLng3 = scanFromRes3.getResults()[0].getGeometry().getLocation();
			final LatLng relLatLng4 = scanFromRes4.getResults()[0].getGeometry().getLocation();
			final LatLng relLatLng5 = telHai;
			
			Collections.sort(cityByDistance, new Comparator<LocationEntry>() {
				public int compare(LocationEntry loc1, LocationEntry loc2) {
					GeocoderResult res1 = GeocodeDataManagerCollection.getAddressGeodata(getEntryFormattedAdress(loc1, null), false, true);
					GeocoderResult res2 = GeocodeDataManagerCollection.getAddressGeodata(getEntryFormattedAdress(loc2, null), false, true);

					double val = Math.min(distanceFromPoint(res1, relLatLng1),
							Math.min(distanceFromPoint(res1, relLatLng2),
									Math.min(distanceFromPoint(res1, relLatLng3),
											Math.min(distanceFromPoint(res1, relLatLng4),
																	distanceFromPoint(res1, relLatLng5))))) -
							Math.min(distanceFromPoint(res2, relLatLng1),
									Math.min(distanceFromPoint(res2, relLatLng2),
											Math.min(distanceFromPoint(res2, relLatLng3),
													Math.min(distanceFromPoint(res2, relLatLng4),
																	distanceFromPoint(res2, relLatLng5)))));
					if (val < 0)
						return -1;
					else if (val > 0)
						return 1;
					else
						return loc1.getCityName().compareTo(loc2.getCityName());
				}

				private double distanceFromPoint(GeocoderResult res, LatLng relLatLng) {
					double lat = res.getGeometry().getLocation().getLat().doubleValue();
					double lng = res.getGeometry().getLocation().getLng().doubleValue();
					double distanceSqr = Math.pow(lat - relLatLng.getLat().doubleValue(), 2) + 
							Math.pow(lng - relLatLng.getLng().doubleValue(), 2);
					return distanceSqr;
				}

			});
		}
	}

	private static void scanStreetInfo() {
		// int i = 0;
		for (int i = 0; i < cityByDistance.size() && GeocodeDataManager.isQueryingAllowed(); ++i) {
			LocationEntry cityEnt = cityByDistance.get(i);
			scanStreetInfo(cityEnt);
			GeocoderResult res =  GeocodeDataManagerCollection.getAddressGeodata(getEntryFormattedAdress(cityEnt, null), false, true);
			if ((res != null) && (res.getFormattedAddress().equals("מודיעין מכבים רעות, ישראל")))
				break;
		}
	}

	private static void scanStreetInfo(LocationEntry cityEnt) {
		int readRow = 1;
		boolean userBreak = false;
		GeocodeDataManager parentMan = GeocodeDataManagerCollection.getGeocodeManager(getEntryFormattedAdress(cityEnt, null));

		while (!XLSUtil.isEndRow(streetWB, readRow) && GeocodeDataManager.isQueryingAllowed() && !userBreak) {
			LocationEntry loc = new LocationEntry(streetWB, readRow);
			++readRow;

			if (cityEnt.isSameCity(loc) && !loc.isCity()) {
				GeocodeDataManagerCollection.getAddressGeodata(getEntryFormattedAdress(loc, parentMan), true, true);
			}
		}
	}
	
	public static String getEntryFormattedAdress(LocationEntry loc, GeocodeDataManager man) {
		if (loc.isCity()) {
			if (loc.getCityName().contains("(") || loc.getCityName().contains(")"))
				return null;
			else {
				String name = loc.getQueryCityName();
				name = name.replaceAll(" - ", " ");
				return name + ", " + DEFAULT_COUNTRY_NAME;
			}
		} else {
			String areaName = loc.getQueryAreaName();
			if ((areaName.matches("רח \\d+")) || (areaName.matches("רח .*")))
				return null;

			areaName.trim();
			if (areaName.matches(".* \\p{InHebrew}"))
				areaName = areaName + "'";
			
			String neighborhoodSuffix = "שכונת ";
			if (areaName.matches("(ש|שכ|שכונה) (\\p{InHebrew}'|י\"\\p{InHebrew})"))
				areaName = areaName.replaceAll("(ש|שכ|שכונה) (\\p{InHebrew}'|י\"\\p{InHebrew})", "שכונה $2");
			else if (areaName.equals("גורדון") || areaName.equals("גורדון א ד"))
				areaName = "א.ד. גורדון";
			else
			{
				areaName = areaName.replaceAll(" סמ\\d+", "");
				areaName = areaName.replaceAll("( \\d+)*$", "");
				areaName = areaName.replaceAll("^שד ", "שדרות ");
				areaName = areaName.replaceAll("^שכ ", neighborhoodSuffix);
				areaName = areaName.replaceAll("^ש ", neighborhoodSuffix);
				areaName = areaName.replaceAll("^סמ ", "סמטת ");
				areaName = areaName.replaceAll("^שכוני ", "שיכוני ");
				areaName = areaName.replaceAll("\\bאסוולדו\\b", "אוסוולדו");
				if (man.getBaseLocationName().contains("באר שבע"))
					areaName = areaName.replaceAll("\\bאוסוולדו\\b", "אוסבלדו");
			}
			
				
			if ((areaName == null) || (areaName.isEmpty()))
				return null;

			return areaName + ", " + man.getBaseLocationName();
		}
	}
}
