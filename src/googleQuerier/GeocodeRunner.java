package googleQuerier;

import googleQuerier.GeocodeDataManager.QueryData;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.code.geocoder.model.GeocoderResult;

import utility.XLSUtil;

public class GeocodeRunner {

	private static final String GEOCODE_OUTPUT_DIR = "GeocodeData";
	private static final String GEOCODE_FILE_PREFIX = "Geo-1";
	private static final String GEOCODE_FILE_POSTFIX = ".dat";
	private static final String DEFAULT_COUNTRY_NAME = ", ישראל";
	private static GeocodeDataManager countryData = null;
	private static GeocodeDataManager currentCityData = null;
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

			File cityFile = new File(workingDir, GEOCODE_FILE_PREFIX + GEOCODE_FILE_POSTFIX);
			countryData = new GeocodeDataManager(cityFile);
			
			scanCityInfo();
			scanStreetInfo();
			//gatherCityInfo();
			// scanCityStreets();
		}
		if (countryData != null)
			countryData.close();
		if (currentCityData != null)
			currentCityData.close();
	}

	private static void scanCityInfo() {
					
		cityByDistance = new ArrayList<LocationEntry>();
		int readRow = 1;
		boolean userBreak = false;
		while (!XLSUtil.isEndRow(streetWB, readRow) && !userBreak) {
			LocationEntry loc = new LocationEntry(streetWB, readRow);
			++readRow;
			
			loc = loc.getCityLocation();
			if (!cityByDistance.contains(loc))
			{
				GeocoderResult res = getAddressGeodata(loc);
				if (res != null)
				{
					cityByDistance.add(loc);
				}
			}
			
		}
		
		QueryData scanFromRes = countryData.getOrCreateQueryData("תל אביב, ישראל");
		if ((scanFromRes != null) && (scanFromRes.getDescription().isRecordExists()))
		{
			final double relLat = scanFromRes.getResult().getGeometry().getLocation().getLat().doubleValue();
			final double relLng = scanFromRes.getResult().getGeometry().getLocation().getLng().doubleValue();
			
			Collections.sort(cityByDistance, new Comparator<LocationEntry>() {
			    public int compare(LocationEntry loc1, LocationEntry loc2) {
			    	GeocoderResult res1 = getAddressGeodata(loc1);
			    	GeocoderResult res2 = getAddressGeodata(loc2);
			    	
			    	double val = distanceFromPoint(res1, relLat, relLng) - distanceFromPoint(res2, relLat, relLng);
			    	if (val < 0) return -1;
			    	else if (val > 0) return 1;
			    	else return 0;
			    }

				private double distanceFromPoint(GeocoderResult res, double relLat, double relLng) {
					double lat = res.getGeometry().getLocation().getLat().doubleValue();
					double lng = res.getGeometry().getLocation().getLng().doubleValue();
					double distanceSqr = Math.pow(lat - relLat, 2) + Math.pow(lng - relLng, 2);
					return distanceSqr;
				}
			    
			});
		}
		cityByDistance = cityByDistance;
	}
	
	private static void scanStreetInfo() {
		// TODO Auto-generated method stub
		
	}


	/*private static void gatherCityInfo() throws ClassNotFoundException, IOException {
		int readRow = 1;
		boolean userBreak = false;
		while (!XLSUtil.isEndRow(streetWB, readRow) && !userBreak) {
			LocationEntry loc = new LocationEntry(streetWB, readRow);
			++readRow;
			getAddressGeodata(loc.getCityLocation());
		}

	}*/

	/**
	 * @param loc
	 * @return
	 */
	private static GeocoderResult getAddressGeodata(LocationEntry loc) {
		ArrayList<String> addressNames = getAddressSuggestions(loc);
		if (addressNames != null) {
			for (int i = 0; i < addressNames.size(); ++i) {
				QueryData qd = countryData.getOrCreateQueryData(addressNames.get(i));
				if ((qd != null) && (qd.getDescription().isRecordExists()) && (!qd.getDescription().isPartialMatch()))
					return qd.getResult();
			}
		}
		return null;
	}

	private static ArrayList<String> getAddressSuggestions(LocationEntry loc) {
		if (loc.isCity()) {
			if (loc.getCityName().contains("(") || loc.getCityName().contains(")"))
				return null;
			else {
				ArrayList<String> ret = new ArrayList<String>();
				ret.add(loc.getQueryCityName() + DEFAULT_COUNTRY_NAME);
				return ret;
			}
		}
		else
			return null;
	}

}
