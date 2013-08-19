package googleQuerier;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import googleQuerier.GeocodeDataManager.QueryData;

import com.google.code.geocoder.model.GeocoderResult;
import com.pairapp.utilities.SimpleCache;

public class GeocodeDataManagerCollection {
	
	private static final String DEFAULT_COUNTRY_NAME = "ישראל";
	public static final String DEFAULT_COUNTRY_CODE = "IL";
	private static GeocodeDataManager countryData = null;
	private static SimpleCache<String, GeocodeDataManager> cityDataManagers = null;
	
	static String workingDirectory;
	
	public static GeocodeDataManager getCountryData() { return countryData;}
	
	public static void initialize(String workingDir) 
	{
		workingDirectory = workingDir;
		cityDataManagers = new SimpleCache<String, GeocodeDataManager>(4, false,
			new SimpleCache.Destructor<GeocodeDataManager>() {
				public void destroy(GeocodeDataManager man) {
					man.close();
				}
			});
		
		countryData = new GeocodeDataManager(workingDirectory, DEFAULT_COUNTRY_NAME);
	}
	
	public static GeocodeDataManager getGeocodeManager(String addressName, boolean extractForParent) {
		if (extractForParent)
			addressName = getParentAddressName(addressName);
		return getGeocodeManager(addressName);
	}
	
	public static GeocoderResult getAddressGeodata(String addressName, boolean allowCreate, boolean guessName) {
		if (addressName != null) {
			GeocodeDataManager man = GeocodeDataManagerCollection.getGeocodeManager(addressName, true);
			if (man != null) {
				QueryData qd = null;
				for(int i = 0 ; i < 6; i++)
				{
					String checkName = addressName;
					if (i >= 2)
						checkName = removePrefixName(checkName);
					if ((i == 1) || (i == 3))
						checkName = reversedName(checkName);
					if (i == 5)
						checkName = extractMainNameOnly(checkName);					
					if (checkName != null)
					{
						qd = man.getOrCreateQueryData(checkName, allowCreate);
					
						if (isHasRecord(qd) && (qd.getResults()[0].getFormattedAddress().equals(checkName)))
							return qd.getResults()[0];
						//Check if this is a city name
						if ((addressName.indexOf(",") == addressName.lastIndexOf(",")) ||
								(guessName == false))
							return null;
					}
				}
			}
		}
		return null;
	}

	
	private static String extractMainNameOnly(String checkName) {
		int delimSpace = checkName.indexOf(" ");
		int delimPeriod = checkName.indexOf(",");
		if ((delimSpace < delimPeriod) && (delimSpace != -1))
		{
			checkName = checkName.substring(0, delimSpace) + checkName.substring(delimPeriod);
			if (checkName.length() > 4)
				return checkName;
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

	

	private static String reversedName(String areaName) {
		String postFix = areaName.substring(areaName.indexOf(","));
		areaName = areaName.substring(0,areaName.length() - postFix.length());
		
		Pattern p = Pattern
				.compile("^(?<pref>((בי\"ס|בית|שדרות|בסיס|בעל|בן|בר|כיכר|מרכז רפואי|מרכז|ככר|הקריה|קרית|שכונת|פרופ'|פרופסור|פרופ|מלון|קניון|נאות|סמטת|הרב|גבעת|אבא|דרך|הרבי|נחלת|חטיבת|אזור|חוש|הר|רבי|רמת|רחבת|אלוף|ד\"ר|דוקטור|אבו) )?+)(?<part1>(.+?)) (?<!\\b(א|דיר|נחלת|בן|דה|בר|אל) )(?<part2>.+)$");
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
		addressName = addressName.replaceAll("^דרך ", "");
		addressName = addressName.replaceAll("^אלוף ", "");
		addressName = addressName.replaceAll("^נחלת ", "");
		addressName = addressName.replaceAll("^ככר ", "");
		addressName = addressName.replaceAll("^הרב ", "");
		addressName = addressName.replaceAll("^הרבי ", "");
		addressName = addressName.replaceAll("^ה ", "");
		return addressName;
	}

	public static GeocodeDataManager getGeocodeManager(String addressName) {
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
							retMan = new GeocodeDataManager(workingDirectory, formattedAddress);
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

	public static void deinitialize(String workingDir) {
		if (countryData != null)
			countryData.close();
		if (cityDataManagers != null)
			cityDataManagers.clear();
	}


}
