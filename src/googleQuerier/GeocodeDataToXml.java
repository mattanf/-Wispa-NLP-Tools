package googleQuerier;

import java.io.File;
import java.io.PrintStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeSet;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.google.code.geocoder.model.GeocoderResult;
import com.google.code.geocoder.model.LatLng;
import com.pairapp.engine.parser.location.LocationLanguageData;
import com.pairapp.engine.parser.location.NamedLocation;
import com.pairapp.engine.parser.location.NamedLocation.LocationType;

@SuppressWarnings("unused")
public class GeocodeDataToXml {

	private LocationLanguageData langData;

	public GeocodeDataToXml(LocationLanguageData langData) {
		this.langData = langData;
	}

	public boolean writeXMLtoFile(String parentName, String parentCode, Iterator<NamedLocation> results,
			String workingDir, LocationType parentLocType, Map<String, String> resultToCode) {
		TreeSet<NamedLocation> orderedResults = new TreeSet<NamedLocation>(new Comparator<NamedLocation>() {
			public int compare(NamedLocation res1, NamedLocation res2) {
				int val = res1.getFormattedName().compareTo(res2.getFormattedName());
				if (val == 0)
					val = res1.getType().ordinal() - res2.getType().ordinal();
				if (val == 0)
					val = Double.compare(res1.getLatitude(), res2.getLatitude());
				if (val == 0)
					val = Double.compare(res1.getLongitude(), res2.getLongitude());
				return val;
			}
		});
		while (results.hasNext()) {
			orderedResults.add(results.next());
		}

		try {
			File xmlFilename = new File(workingDir, "LocationData-" + parentCode + ".xml");

			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

			// root elements
			Document doc = docBuilder.newDocument();
			Element rootElement = doc.createElement("LocationData");
			doc.appendChild(rootElement);

			Iterator<Entry<NamedLocation, String[]>> itCountryRegions = null;
			if (resultToCode != null)
				itCountryRegions = countryRegions.entrySet().iterator();
			Iterator<NamedLocation> itRes = orderedResults.iterator();
			String prevAddressName = null;
			LocationType prevLocationType = null;
			while ((itRes.hasNext()) || ((itCountryRegions != null) && (itCountryRegions.hasNext()))) {
				LatLng lngLat = null;
				String regionCode = null;
				NamedLocation key = null;
				if (itRes.hasNext()) {
					key = itRes.next();
				}
				else {
					assert (itCountryRegions != null && itCountryRegions.hasNext());
					key = itCountryRegions.next().getKey();
				}
				LocationType curLocType = key.getType();
				String addressName = key.getMainName();
				String code = key.getCode();
				//jump over duplicates
				if (addressName.equals(prevAddressName) && curLocType.equals(prevLocationType))
					continue;
				
				if (key.hasLatLong()) {
					lngLat = new LatLng(new BigDecimal(key.getLatitude()), new BigDecimal(key.getLongitude()));
				}
				if ((resultToCode != null) && (code == null))
					code = resultToCode.get(key.getFormattedName());
				
				if (parentLocType == LocationType.Country)
					regionCode = getCountryRegionCode(code, lngLat, curLocType);
				
				if (!disqualifyName(addressName)) {
					if (((parentLocType == LocationType.Country) && ((curLocType == LocationType.City) || (curLocType == LocationType.CountryRegion))) ||
							((parentLocType == LocationType.City) && ((curLocType == LocationType.CityRegion) || (curLocType == LocationType.Street)))) {

						Element locElem = doc.createElement(curLocType.name());
						rootElement.appendChild(locElem);
						if (code != null) {
							Element coElem = doc.createElement("Code");
							coElem.setTextContent(code);
							locElem.appendChild(coElem);
						}

						if (regionCode != null) {
							Element coElem = doc.createElement("RegionCode");
							coElem.setTextContent(regionCode);
							locElem.appendChild(coElem);
						}

						Element nnElem = doc.createElement("NativeName");
						nnElem.setTextContent(addressName);
						locElem.appendChild(nnElem);

						Element pnElem = doc.createElement("ParentName");
						pnElem.setTextContent(parentName);
						locElem.appendChild(pnElem);

						String[] alternativeNames = getAlternativeNames(addressName, curLocType, parentName);
						if (alternativeNames != null) {
							for (int i = 0; i < alternativeNames.length; ++i) {
								Element altElem = doc.createElement("AlternateName");
								altElem.setTextContent(alternativeNames[i]);
								locElem.appendChild(altElem);
							}
						}

						if ((lngLat != null)) {
							Element latElem = doc.createElement("Latitude");
							latElem.setTextContent(String.format("%.6f",lngLat.getLat().doubleValue()));
							locElem.appendChild(latElem);

							Element lngElem = doc.createElement("Longitude");
							lngElem.setTextContent(String.format("%.6f",lngLat.getLng().doubleValue()));
							locElem.appendChild(lngElem);
						}
					}
				}
				
				prevAddressName = addressName;
				prevLocationType = curLocType;
				

			}

			// write the content into xml file
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new PrintStream(xmlFilename));
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
			transformer.transform(source, result);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	private static HashMap<NamedLocation, String[]> countryRegions = null;
	static {
		countryRegions = new HashMap<>();
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "1", null, "אזור השרון", null, new Double(
				32.153), new Double(34.863)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "3", null, "אצבע הגליל", null, new Double(
				33.233), new Double(35.577)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "R101", null, "יהודה ושומרון", null, new Double(
				31.78), new Double(35.198)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "R102", null, "הנגב", null, new Double(
				31.25), new Double(34.782)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "2", null, "גבעתיים/רמת גן", null, new Double(
				32.068), new Double(34.816)), new String[] { "8600", "6300" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "4", null, "באר שבע", null, new Double(31.25),
				new Double(34.782)), new String[] { "9000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "5", null, "ירושלים", null, new Double(31.78),
				new Double(35.198)), new String[] { "3000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "6", null, "תל אביב", null,
				new Double(32.080), new Double(34.773)), new String[] { "5000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "7", null, "חיפה", null, new Double(32.772265),
				new Double(35.020294)), new String[] { "4000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "R103", null, "הכרמל", null, new Double(32.772265),
				new Double(35.020294)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "8", null, "חולון, בת ים, ראשון", null, new Double(31.918947),
				new Double(34.812927)), null);//new String[] { "6600", "6200", "8300" });
	}


	private String getCountryRegionCode(String code, LatLng lngLat, LocationType locType) {
		String shortestDistanceCode = null;
		if (locType == LocationType.City || locType == LocationType.CityRegion) {
			double curShortestDistance = Double.MAX_VALUE;
			Iterator<Entry<NamedLocation, String[]>> itEnt = countryRegions.entrySet().iterator();
			while (itEnt.hasNext()) {
				Entry<NamedLocation, String[]> next = itEnt.next();
				if (next.getValue() != null) {
					for (int i = next.getValue().length - 1; i >= 0; --i) {
						if (next.getValue()[i].equals(code))
							return next.getKey().getCode();
					}
				} else {
					double dist = distXY(lngLat.getLat().doubleValue() - next.getKey().getLatitude(), lngLat.getLng()
							.doubleValue() - next.getKey().getLongitude());
					if (dist < curShortestDistance) {
						curShortestDistance = dist;
						shortestDistanceCode = next.getKey().getCode();
					}
				}
			}
			if ((locType == LocationType.CityRegion) && (shortestDistanceCode != null) && (shortestDistanceCode.equals(code)))
			{
				shortestDistanceCode = null;
			}
		}
		return shortestDistanceCode;
	}

	private double distXY(double d1, double d2) {
		return Math.sqrt(d1 * d1 + d2 * d2);
	}

	private boolean disqualifyName(String addressName) {
		addressName = addressName.trim();
		if (addressName.matches("\\d+"))
			return true;
		addressName = addressName.replace("'", "");
		addressName = addressName.replace("\"", "");
		if (addressName.length() < 2)
			return true;
		return false;
	}

	private String[] getAlternativeNames(String name, LocationType type, String parentName) {
		HashMap<String, Integer> arr = new HashMap<>();
		langData.generateNamePermutations(name, arr, parentName);

		arr.remove(name);
		return arr.keySet().toArray(new String[arr.size()]);
	}


}
