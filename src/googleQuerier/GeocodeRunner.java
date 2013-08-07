package googleQuerier;

import googleQuerier.GeocodeDataManager.QueryData;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.code.geocoder.model.GeocoderResult;
import com.pairapp.engine.parser.location.NamedLocation.LocationType;
import com.pairapp.utilities.SimpleCache;

import utility.XLSUtil;

public class GeocodeRunner {

	private static final String GEOCODE_OUTPUT_DIR = "GeocodeData";
	private static final String DEFAULT_COUNTRY_NAME = "ישראל";
	private static GeocodeDataManager countryData = null;
	private static SimpleCache<String, GeocodeDataManager> cityDataManagers = null;
	private static ArrayList<LocationEntry> cityByDistance = null;
	private static XSSFWorkbook streetWB;
	private static String workingDir;

	/**
	 * This program recives an xlsx file that contains a list of location. It queries google's geocode web service for
	 * information, saves the information to the disk, than later updates the xlsx file with the information,
	 * 
	 * @param args
	 * @throws IOException
	 * @throws ClassNotFoundException
	 */
	public static void main(String[] args) throws ClassNotFoundException, IOException {

		boolean gatherDataFromGoogle = false;
		if ((args.length > 0) && (args[0].equals("-scan")))
			gatherDataFromGoogle = true;

		cityDataManagers = new SimpleCache<String, GeocodeDataManager>(4, false,
				new SimpleCache.Destructor<GeocodeDataManager>() {
					public void destroy(GeocodeDataManager man) {
						man.close();
					}
				});

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

			countryData = new GeocodeDataManager(workingDir, DEFAULT_COUNTRY_NAME);

			if (gatherDataFromGoogle) {
				scanCityInfo();
				scanStreetInfo();
			}
			outputGatheredData();
		}
		if (countryData != null)
			countryData.close();
		if (cityDataManagers != null)
			cityDataManagers.clear();
	}

	private static void outputGatheredData() {
		GeocodeDataToXml serializer = new GeocodeDataToXml();

		ArrayList<GeocoderResult> citiesToWrite = new ArrayList<GeocoderResult>();
		Iterator<GeocoderResult> cityIt = countryData.getUniqueGeoResults();
		while (cityIt.hasNext()) {
			GeocoderResult cityRes = cityIt.next();
			GeocodeDataManager man = getGeocodeManager(cityRes.getFormattedAddress());
			if ((man != null) && (man.getUniqueResultCount() > 30)) {
				serializer.writeXMLtoFile(man.getBaseLocationName(), man.getUniqueGeoResults(), workingDir,
						LocationType.City);
				citiesToWrite.add(cityRes);
			}
		}

		serializer.writeXMLtoFile(countryData.getBaseLocationName(), citiesToWrite.iterator(), workingDir,
				LocationType.Country);
	}

	private static void scanCityInfo() {

		cityByDistance = new ArrayList<LocationEntry>();
		int readRow = 1;
		boolean userBreak = false;
		while (!XLSUtil.isEndRow(streetWB, readRow) && !userBreak) {
			LocationEntry loc = new LocationEntry(streetWB, readRow);
			++readRow;

			loc = loc.generateCityLocation();
			if (!cityByDistance.contains(loc)) {
				GeocoderResult res = getAddressGeodata(getEntryFormattedAdress(loc, null), true);
				if (res != null) {
					cityByDistance.add(loc);
				}
			}

		}

		QueryData scanFromRes1 = countryData.getOrCreateQueryData("תל אביב, ישראל");
		QueryData scanFromRes2 = countryData.getOrCreateQueryData("ירושליים, ישראל");
		QueryData scanFromRes3 = countryData.getOrCreateQueryData("באר שבע, ישראל");
		//scanFromRes3 = scanFromRes2 = scanFromRes1;
		if ((scanFromRes1 != null) && (scanFromRes1.getDescription().isRecordExists())) {
			final double relLat1 = scanFromRes1.getResults()[0].getGeometry().getLocation().getLat().doubleValue();
			final double relLng1 = scanFromRes1.getResults()[0].getGeometry().getLocation().getLng().doubleValue();
			final double relLat2 = scanFromRes2.getResults()[0].getGeometry().getLocation().getLat().doubleValue();
			final double relLng2 = scanFromRes2.getResults()[0].getGeometry().getLocation().getLng().doubleValue();
			final double relLat3 = scanFromRes3.getResults()[0].getGeometry().getLocation().getLat().doubleValue();
			final double relLng3 = scanFromRes3.getResults()[0].getGeometry().getLocation().getLng().doubleValue();

			Collections.sort(cityByDistance, new Comparator<LocationEntry>() {
				public int compare(LocationEntry loc1, LocationEntry loc2) {
					GeocoderResult res1 = getAddressGeodata(getEntryFormattedAdress(loc1, null), false);
					GeocoderResult res2 = getAddressGeodata(getEntryFormattedAdress(loc2, null), false);

					double val = Math.min(
							distanceFromPoint(res1, relLat1, relLng1),
							Math.min(distanceFromPoint(res1, relLat2, relLng2),
									distanceFromPoint(res1, relLat3, relLng3))) -
							Math.min(
									distanceFromPoint(res2, relLat1, relLng1),
									Math.min(distanceFromPoint(res2, relLat2, relLng2),
											distanceFromPoint(res2, relLat3, relLng3)));
					if (val < 0)
						return -1;
					else if (val > 0)
						return 1;
					else
						return loc1.getCityName().compareTo(loc2.getCityName());
				}

				private double distanceFromPoint(GeocoderResult res, double relLat, double relLng) {
					double lat = res.getGeometry().getLocation().getLat().doubleValue();
					double lng = res.getGeometry().getLocation().getLng().doubleValue();
					double distanceSqr = Math.pow(lat - relLat, 2) + Math.pow(lng - relLng, 2);
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

		}
	}

	private static void scanStreetInfo(LocationEntry cityEnt) {
		int readRow = 1;
		boolean userBreak = false;
		GeocodeDataManager parentMan = getGeocodeManager(getEntryFormattedAdress(cityEnt, null));

		while (!XLSUtil.isEndRow(streetWB, readRow) && GeocodeDataManager.isQueryingAllowed() && !userBreak) {
			LocationEntry loc = new LocationEntry(streetWB, readRow);
			++readRow;

			if (cityEnt.isSameCity(loc) && !loc.isCity()) {
				GeocoderResult res = getAddressGeodata(getEntryFormattedAdress(loc, parentMan), true);
			}
		}
	}

	/**
	 * 
	 * @param loc
	 * @return
	 */
	private static GeocoderResult getAddressGeodata(String addressName, boolean allowCreate) {
		if (addressName != null) {
			GeocodeDataManager man = getGeocodeManager(getParentAddressName(addressName));
			if (man != null) {
				QueryData qd = null;
				for(int i = 0 ; i < 4 ; i++)
				{
					String checkName = addressName;
					if (i >= 2)
						checkName = removePrefixName(checkName);
					if (i % 2 == 1)
						checkName = reversedName(checkName);
					if (checkName != null)
					{
						qd = man.getOrCreateQueryData(checkName, allowCreate);
					
						if (isHasRecord(qd))
							return qd.getResults()[0];
						//Check if this is a city name
						if (addressName.indexOf(",") == addressName.lastIndexOf(","))
							return null;
					}
				}
			}
		}
		return null;
	}

	
	private static GeocodeDataManager getGeocodeManager(String addressName) {
		GeocodeDataManager retMan = null;
		if ((addressName != null) && (addressName.isEmpty() == false)) {
			if (addressName.equals(DEFAULT_COUNTRY_NAME))
				retMan = countryData;
			else {
				String parentName = getParentAddressName(addressName);
				GeocodeDataManager parentMan = getGeocodeManager(parentName);
				if (parentMan != null) {
					QueryData qd = parentMan.getOrCreateQueryData(addressName, false);
					String formattedAddress = ((qd != null) && (qd.getResults() != null)) ? qd.getResults()[0]
							.getFormattedAddress() : null;
					if (formattedAddress != null) {
						retMan = cityDataManagers.get(formattedAddress);
						if (retMan == null) {
							retMan = new GeocodeDataManager(workingDir, formattedAddress);
							cityDataManagers.put(formattedAddress, retMan);
						}
					}
				}
			}
		}
		return retMan;
	}

	public static String getParentAddressName(String addressName) {
		if (addressName != null) {
			int delim = addressName.indexOf(",");
			if (delim != -1)
				return addressName.substring(delim + 1).trim();
		}
		return null;

	}

	/**
	 * @param qd
	 * @return
	 */
	private static boolean isHasRecord(QueryData qd) {
		if ((qd != null) && (qd.getDescription().isRecordExists())) {
			String curType = qd.getResults()[0].getTypes().size() > 0 ?  qd.getResults()[0].getTypes().get(0) : null;
			if (curType != null) {
				return ((curType.equals("locality")) || (curType.equals("route")) || (curType.equals("neighborhood")) ||
						(curType.equals("park")) || (curType.equals("point_of_interest")));
			}
		}
		return false;
	}

	private static String getEntryFormattedAdress(LocationEntry loc, GeocodeDataManager man) {
		if (loc.isCity()) {
			if (loc.getCityName().contains("(") || loc.getCityName().contains(")"))
				return null;
			else {
				return loc.getQueryCityName() + ", " + DEFAULT_COUNTRY_NAME;
			}
		} else {
			String areaName = loc.getQueryAreaName();
			if ((areaName.matches("רח \\d+")) || (areaName.matches("רח .*")))
				return null;

			areaName.trim();
			areaName = areaName.replaceAll("(( \\d+| \\p{InHebrew}))*$", "");
			areaName = areaName.replaceAll(" סמ\\d+", "");
			areaName = areaName.replaceAll("^שד ", "שדרות ");
			areaName = areaName.replaceAll("^שכ ", "שכונת ");
			areaName = areaName.replaceAll("^ש ", "שכונת ");
			areaName = areaName.replaceAll("^סמ ", "סמטת ");
			areaName = areaName.replaceAll("\\bאסוולדו\\b", "אוסוולדו");
			if (man.getBaseLocationName().contains("באר שבע"))
				areaName = areaName.replaceAll("\\bאוסוולדו\\b", "אוסבלדו");
				
			if ((areaName == null) || (areaName.isEmpty()))
				return null;

			return areaName + ", " + man.getBaseLocationName();
		}
	}

	private static String reversedName(String areaName) {
		String postFix = areaName.substring(areaName.indexOf(","));
		areaName = areaName.substring(0,areaName.length() - postFix.length());
		
		Pattern p = Pattern
				.compile("^(?<pref>((בי\"ס|בית|שדרות|בסיס|בעל|בן|בר|כיכר|מרכז רפואי|מרכז|ככר|הקריה|קרית|שכונת|פרופ'|פרופסור|פרופ|מלון|קניון|נאות|סמטת|הרב|גבעת|אבא|דרך|הרבי|חטיבת|אזור|חוש|הר|רבי|רחבת|אלוף|ד\"ר|דוקטור|אבו) )?+)(?<part1>(.+?)) (?<!\\b(א|דיר|בן|דה|בר|אל) )(?<part2>.+)$");
		Matcher mch = p.matcher(areaName);
		if (mch.find()) {
			String s = mch.group("pref") + mch.group("part2") + " " + mch.group("part1");
			return s + postFix;
		}
		return null;
	}
	
	private static String removePrefixName(String addressName) {
		addressName = addressName.replaceAll("^שדרות ", "");
		addressName = addressName.replaceAll("^שכונת ", "");
		addressName = addressName.replaceAll("^סמטת ", "");
		addressName = addressName.replaceAll("^קרית ", "");
		return addressName;
	}


}
