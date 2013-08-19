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

	public boolean writeXMLtoFile(String parentName, String parentCode, Iterator<GeocoderResult> results,
			String workingDir, LocationType parentLocType, Map<String, String> resultToCode) {
		TreeSet<GeocoderResult> orderedResults = new TreeSet<GeocoderResult>(new Comparator<GeocoderResult>() {
			public int compare(GeocoderResult res1, GeocoderResult res2) {
				return res1.getFormattedAddress().compareTo(res2.getFormattedAddress());
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
			Iterator<GeocoderResult> itRes = orderedResults.iterator();
			while ((itRes.hasNext()) || ((itCountryRegions != null) && (itCountryRegions.hasNext()))) {
				LocationType curLocType;
				String addressName;
				String code = null;
				LatLng lngLat = null;
				String regionCode = null;
				if (itRes.hasNext()) {
					GeocoderResult res = itRes.next();
					curLocType = getLocationType(res);
					addressName = res.getAddressComponents().get(0).getLongName();

					if (resultToCode != null)
						code = resultToCode.get(res.getFormattedAddress());

					if ((res.getGeometry() != null) && (res.getGeometry().getLocation() != null) &&
							(res.getGeometry().getLocation().getLat() != null)) {
						lngLat = res.getGeometry().getLocation();
					}
					 if (parentLocType == LocationType.Country)
						 regionCode = getCountryRegionCode(code, lngLat, curLocType);
				} else {
					assert (itCountryRegions != null && itCountryRegions.hasNext());
					NamedLocation key = itCountryRegions.next().getKey();

					curLocType = LocationType.CountryRegion;
					addressName = key.getMainName();
					code = key.getCode();

					if (key.hasLatLong()) {
						lngLat = new LatLng(new BigDecimal(key.getLatitude()), new BigDecimal(key.getLongitude()));
					}
					if (parentLocType == LocationType.Country)
						regionCode = getCountryRegionCode(code, lngLat, curLocType);
				}

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
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "2", null, "אצבע הגליל", null, new Double(
				33.233), new Double(35.577)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "3", null, "יהודה ושומרון", null, new Double(
				31.78), new Double(35.198)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "5", null, "הנגב", null, new Double(
				31.25), new Double(34.782)), null);
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "6", null, "גבעתיים/רמת גן", null, new Double(
				32.068), new Double(34.816)), new String[] { "8600", "6300" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "7", null, "באר שבע", null, new Double(31.25),
				new Double(34.782)), new String[] { "9000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "8", null, "ירושלים", null, new Double(31.78),
				new Double(35.198)), new String[] { "3000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "9", null, "תל אביב", null,
				new Double(32.080), new Double(34.773)), new String[] { "5000" });
		countryRegions.put(new NamedLocation(LocationType.CountryRegion, "10", null, "חיפה", null, new Double(32.011),
				new Double(34.784)), new String[] { "4000" });
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

	private LocationType getLocationType(GeocoderResult res) {
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

}
