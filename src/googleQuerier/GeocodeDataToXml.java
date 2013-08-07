package googleQuerier;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.google.code.geocoder.model.GeocoderResult;
import com.pairapp.engine.parser.location.NamedLocation.LocationType;
import com.pairapp.utilities.StringUtilities;

public class GeocodeDataToXml {

	public boolean writeXMLtoFile(String parentName, Iterator<GeocoderResult> results, String workingDir,
			LocationType parentLocType) {
		try {
			File xmlFilename = new File(workingDir, "LocationData-" +
					StringUtilities.removeUnsafeFileChars(parentName) + ".xml");

			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

			// root elements
			Document doc = docBuilder.newDocument();
			Element rootElement = doc.createElement("LocationData");
			doc.appendChild(rootElement);

			Iterator<GeocoderResult> itRes = results;
			while (itRes.hasNext()) {
				GeocoderResult res = itRes.next();
				LocationType curLocType = getLocationType(res);
				String addressName = res.getAddressComponents().get(0).getLongName();
				// String shortName = res.getAddressComponents().get(0).getShortName();
				// if (shortName.equals(addressName))
				// shortName = null;
				if (!disqualifyName(addressName)) {
					if (((parentLocType == LocationType.Country) && ((curLocType == LocationType.City) || (curLocType == LocationType.CountryRegion))) ||
							((parentLocType == LocationType.City) && ((curLocType == LocationType.CityRegion) || (curLocType == LocationType.Street)))) {
						Element locElem = doc.createElement(curLocType.name());
						rootElement.appendChild(locElem);

						if (curLocType == LocationType.City) {
							String code = getCode(addressName);
							if (code != null) {
								Element coElem = doc.createElement("Code");
								coElem.setTextContent(code);
								locElem.appendChild(coElem);
							}
						}

						Element nnElem = doc.createElement("NativeName");
						nnElem.setTextContent(addressName);
						locElem.appendChild(nnElem);

						Element pnElem = doc.createElement("ParentName");
						pnElem.setTextContent(parentName);
						locElem.appendChild(pnElem);

						String[] alternativeNames = getAlternativeNames(addressName, curLocType);
						if (alternativeNames != null) {
							for (int i = 0; i < alternativeNames.length; ++i) {
								Element altElem = doc.createElement("AlternateName");
								altElem.setTextContent(alternativeNames[i]);
								locElem.appendChild(altElem);
								// if (alternativeNames[i].equals(shortName))
								// shortName = null;
							}
						}
						/*
						 * if ((shortName != null) && (!shortName.matches("\\d*"))) { Element altElem =
						 * doc.createElement("AlternateName"); altElem.setTextContent(shortName);
						 * locElem.appendChild(altElem); }
						 */

						if ((res.getGeometry() != null) && (res.getGeometry().getLocation() != null) &&
								(res.getGeometry().getLocation().getLat() != null)) {
							Element latElem = doc.createElement("Latitude");
							latElem.setTextContent(res.getGeometry().getLocation().getLat().toString());
							locElem.appendChild(latElem);

							Element lngElem = doc.createElement("Longitude");
							lngElem.setTextContent(res.getGeometry().getLocation().getLng().toString());
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
			transformer.transform(source, result);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	private String getCode(String name) {
		if (name.startsWith("תל אביב"))
			return "5000";
		if (name.equals("ירושליים") || name.equals("ירושלים"))
			return "3000";
		if (name.equals("באר שבע"))
			return "9000";
		if (name.equals("פתח תקווה"))
			return "7900";
		return null;
	}

	private boolean disqualifyName(String addressName) {
		addressName = addressName.trim();
		if (addressName.matches("\\d+"))
			return true;
		addressName = addressName.replace("'", "");
		addressName = addressName.replace("\"", "");
		if (addressName.length() < 2)
			return true;
		if (addressName.equals("אזור"))
			return true;
		return false;
	}

	private String[] getAlternativeNames(String name, LocationType type) {
		if (type == LocationType.City) {
			if (name.equals("באר שבע"))
				return new String[] { "ב\"ש" };
			else if (name.equals("תל אביב יפו"))
				return new String[] { "ת\"א", "תל אביב", "יפו", "תל אביב יפו" };
			else if (name.equals("פתח תקווה"))
				return new String[] { "פ\"ת" };
			return null;
		}
		if (name.contains("דיזנגוף"))
			return new String[] { name.replaceAll("דיזנגוף", "דיזינגוף") };
		if (name.contains("נחום סוקולוב"))
			return new String[] { name.replaceAll("נחום סוקולוב", "סוקולוב") };
		if (name.contains("אוסוולדו"))
			return new String[] { name.replaceAll("אוסוולדו", "אוסבלדו") };
		if (name.contains("אוסבלדו"))
			return new String[] { name.replaceAll("אוסבלדו", "אוסוולדו") };
		
		return null;

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
			else
				return null;
		}
		return null;
	}

}
