import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map.Entry;
import java.util.TreeSet;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utility.XLSUtil;

public class ConvertIsraelStreetDBtoXML {

	static final int ROW_CITY_ID = 2;
	static final int ROW_CITY_NAME = 3;
	static final int ROW_STREET_ID = 4;
	static final int ROW_STREET_NAME = 5;

	static final int STREET_COUNT_DUMP = 10;
	static final int STREET_COUNT_TO_DEFAULT = 30;

	static FileWriter countryOutput = null;
	static FileWriter defaultCityOutput = null;
	static String outputDir = "E:\\MyProjects\\FBAds-Server\\war\\ReferenceData\\LocationData\\LocationData-1\\";

	/**
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException {

		XSSFWorkbook wb = XLSUtil.openXLS("C:\\Public\\Train Data\\rehovot.xlsx");

		countryOutput = new FileWriter(outputDir + "LocationData-1.xml");
		defaultCityOutput = new FileWriter(outputDir + "LocationData-1-X.xml");
		countryOutput.write("<LocationData>\n\r");
		defaultCityOutput.write("<LocationData>\n\r");

		HashMap<String, String> hs1 = new HashMap<>();
		HashMap<String, String> hs2 = new HashMap<>();
		// HashSet<String> hs3 = new HashSet<>();
		// HashSet<String> hs4 = new HashSet<>();
		HashMap<String, String> streetData = new HashMap<>();
		HashMap<String, String> neighborhoodData = new HashMap<>();
		String prevCityId = null;
		String prevCityName = null;
		int readRow = 1;
		while (!XLSUtil.isEndRow(wb, readRow)) {
			++readRow;
			String cityId = XLSUtil.getCellString(wb, readRow, ROW_CITY_ID);
			String cityName = XLSUtil.getCellString(wb, readRow, ROW_CITY_NAME);
			String streetId = XLSUtil.getCellString(wb, readRow, ROW_STREET_ID);
			String streetName = XLSUtil.getCellString(wb, readRow, ROW_STREET_NAME);

			if (streetName.equals(cityName))
				continue;

			if (cityName.indexOf(")") != -1) {
				hs1.put(cityName.substring(cityName.indexOf(")")), cityName);
			}
			if (streetName.indexOf(")") != -1) {
				hs2.put(streetName.substring(streetName.indexOf(")")), cityName);
			}

			cityName = reformantCityName(cityName);
			if (!cityId.equals(prevCityId)) {
				saveCityData(prevCityId, prevCityName, streetData, neighborhoodData);
				streetData = new HashMap<>();
				neighborhoodData = new HashMap<>();
			}

			if (cityName != null) {
				reformantStreetName(streetId, streetName, streetData, neighborhoodData);
			}

			prevCityId = cityId;
			prevCityName = cityName;

		}

		countryOutput.write("</LocationData>");
		defaultCityOutput.write("</LocationData>");
		countryOutput.close();
		defaultCityOutput.close();
		System.out.println("done.");

		// if (false)
		// saveSpecialLists(hs1, hs2);
	}

	private static void saveCityData(String cityId, String cityName, HashMap<String, String> streetData,
			HashMap<String, String> neighborhoodData) throws IOException {
		if (streetData.size() > STREET_COUNT_DUMP) {
			String alternativeNames[] = getAlternateCityNames(cityId);
			if (cityName != null) {
				//
				// Write the city data
				//

				countryOutput.write("\t<City>\n\r");
				countryOutput.write("\t<ParentId>1</ParentId>\n\r");
				countryOutput.write("\t\t<Id>" + cityId + "</Id>\n\r");
				countryOutput.write("\t\t<NativeName>" + cityName + "</NativeName>\n\r");
				if (alternativeNames != null) {
					for (String name : alternativeNames)
						countryOutput.write("\t\t<AlternateName>" + name + "</AlternateName>\n\r");
				}
				countryOutput.write("\t</City>\n\r");

				//
				// Write the street data
				//
				FileWriter curCityOutput = defaultCityOutput;
				boolean inSeperateFile = streetData.size() > STREET_COUNT_TO_DEFAULT;
				if (inSeperateFile) {
					curCityOutput = new FileWriter(outputDir + "\\LocationData-1-" + cityId + ".xml");
					curCityOutput.write("<LocationData>\n\r");
				}

				writeStreetData(curCityOutput, cityId, streetData, false);
				writeStreetData(curCityOutput, cityId, neighborhoodData, true);

				if (inSeperateFile) {
					curCityOutput.write("</LocationData>");
					curCityOutput.close();
				}
			}
		}
	}

	private static void writeStreetData(FileWriter output, String cityId, HashMap<String, String> data,
			boolean isNeighborhood) throws IOException {
		Iterator<Entry<String, String>> it = data.entrySet().iterator();
		while (it.hasNext()) {
			Entry<String, String> ent = it.next();
			String prefix = isNeighborhood ? "CityRegion" : "Street";
			String streetNames[] = genFormattedStreetNames(ent.getValue(), isNeighborhood);
			output.write("\t<" + prefix + ">\n\r");
			output.write("\t\t<ParentId>" + cityId + "</ParentId>\n\r");
			output.write("\t\t<Id>" + ent.getKey() + "</Id>\n\r");
			output.write("\t\t<NativeName>" + streetNames[0] + "</NativeName>\n\r");
			for (int i = 1; i < streetNames.length; ++i) {
				output.write("\t\t<AlternateName>" + streetNames[i] + "</AlternateName>\n\r");
			}

			output.write("\t</" + prefix + ">\n\r");
		}
	}

	private static void reformantStreetName(String streetId, String streetName, HashMap<String, String> streetData,
			HashMap<String, String> neighborhoodData) {

		streetName = streetName.trim();
		if (streetName.lastIndexOf(")") != -1)
			streetName = null;
		if ((streetName != null) && (streetName.matches("(א|רח|רח'|רחוב|ככר) .*")))
			streetName = null;
		if ((streetName != null) && (streetName.matches(".*\\d+")))
			streetName = null;
		if (streetName != null) {
			boolean isNeighboorhood = false;
			if (streetName.matches("(ש|שכ|שכונה|שכונת) .*")) {
				streetName = streetName.substring(streetName.indexOf(' ') + 1);
				isNeighboorhood = true;
			} else if (streetName.matches("(קניון|קרית) .*") || streetName.matches("(סנטר) .*"))
				isNeighboorhood = true;
			streetName.trim();
			if (!streetName.isEmpty()) {
				if (isNeighboorhood) {
					neighborhoodData.put(streetId, streetName);
				} else {
					streetData.put(streetId, streetName);
				}
			}
		}
	}

	private static String reformantCityName(String cityName) {
		if (cityName.endsWith(")יישוב(") || cityName.endsWith(")קיבוץ(") || cityName.endsWith(")מושב(") ||
				cityName.endsWith(")קיבוץ(") || cityName.endsWith(")קיבוץ(")) {
			cityName = cityName.substring(0, cityName.lastIndexOf(")"));
		}
		cityName = cityName.trim();
		if (cityName.lastIndexOf(")") != -1)
			cityName = null;
		if ((cityName != null) && (cityName.equals("אזור")))
			cityName = null;
		return cityName;
	}

	private static String[] getAlternateCityNames(String cityId) {
		if (cityId.equals("9000"))
			return new String[] { "ב\"ש" };
		else if (cityId.equals("5000"))
			return new String[] { "ת\"א", "תל אביב", "יפו", "תל אביב יפו" };
		else if (cityId.equals("3000"))
			return new String[] { "ירושליים" };
		else if (cityId.equals("7900"))
			return new String[] { "פ\"ת" };
		return null;
	}

	private static String[] genFormattedStreetNames(String streetName, boolean isNeighborhood) {
		ArrayList<String> streetNameList = new ArrayList<String>();
		streetNameList.add(streetName);
		streetNameList = generateReplacments(streetNameList, new String[] { "סוקולוב", "סוקולוב", "סוקלוב" } );
		streetNameList = generateReplacments(streetNameList, new String[] { "דיזנגוף", "דיזינגוף" });
		streetNameList = generateReplacments(streetNameList, new String[] { "ארלוזורוב", "ארלוזרוב", "ארלזרוב" });
		streetNameList = generateReplacments(streetNameList, new String[] { "פתח תקווה", "פתח תקוה" });
		streetNameList = generateReplacments(streetNameList, new String[] { "גבעתיים", "גבעתים" });
		streetNameList = generateReplacments(streetNameList, new String[] { "גבעתיים", "גבעתים" });
		streetNameList = generateReplacments(streetNameList, new String[] { "קריית", "קרית" });
		return streetNameList.toArray(new String[streetNameList.size()]);
	}

	// Generates
	private static ArrayList<String> generateReplacments(ArrayList<String> streetNameList, String[] names) {

		String regExStr = "\\b((" + names[0];
		for (int i = 1; i < names.length; ++i) {
			regExStr = regExStr + ")|(" + names[i];
		}
		regExStr = regExStr + "))\\b";
		if (streetNameList.get(0).matches(".*" + regExStr + ".*"))
		{
			ArrayList<String> newStreetNameList = new ArrayList<String>();
			Iterator<String> itStreets = streetNameList.iterator();
			while (itStreets.hasNext()) {
				String inpName = itStreets.next();
				for (int j = 0; j < names.length; ++j) {
					String replaceString = inpName.replaceAll(regExStr, names[j]);
					if (!newStreetNameList.contains(replaceString))
						newStreetNameList.add(replaceString);
				}
			}
			return newStreetNameList;
		}
		return streetNameList;
	}

	
	

	/**
	 * @param hs1
	 * @param hs2
	 * @throws IOException
	 */
	/*
	 * private static void saveSpecialLists(HashMap<String, String> hs1, HashMap<String, String> hs2) throws IOException
	 * { FileWriter out1 = new FileWriter("c:\\special1.txt"); FileWriter out2 = new FileWriter("c:\\special2.txt");
	 * Iterator<Entry<String, String>> it1 = hs1.entrySet().iterator(); while (it1.hasNext()) { Entry<String, String> n
	 * = it1.next(); out1.write(n.getKey() + "\t\t" + n.getValue() + "\n\r"); } Iterator<Entry<String, String>> it2 =
	 * hs2.entrySet().iterator(); while (it2.hasNext()) { Entry<String, String> n = it2.next(); out2.write(n.getKey() +
	 * "\t\t" + n.getValue() + "\n\r"); } out1.close(); out2.close(); }
	 */

}
