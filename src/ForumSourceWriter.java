import java.io.File;
import java.io.PrintStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.PostFieldType.Persistency;
import utility.XLSUtil;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

public class ForumSourceWriter {

	final static String COLUMN_ORDER = "Order";
	final static String COLUMN_ID = "Id";
	final static String COLUMN_BILLBOARD = "Billboard";
	final static String COLUMN_PRIVACY = "Privacy";
	final static String COLUMN_NAME = "Name";
	final static String COLUMN_LINK = "Link";
	final static String COLUMN_OWNER = "Owner";
	final static String COLUMN_DESCRIPTION = "Description";
	final static String COLUMN_ENABLED = "Enabled";

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
		int mainHeaderRow = XLSUtil.getHeaderRow(wb, mainSheet, mainHeader, COLUMN_ID, COLUMN_BILLBOARD,
				COLUMN_PRIVACY, COLUMN_NAME);

		if (mainHeaderRow == -1) {
			System.out.println("No proper header found.");
		} else {
			String sourceFilePath = "ForumSources.xml";
			if (inputFile.getParent() != null)
				sourceFilePath = inputFile.getParent() + "\\" + sourceFilePath;
			exportToXML(wb, mainSheet, mainHeader, mainHeaderRow, new File(sourceFilePath));
		}
	}

	private static void exportToXML(final Workbook wb, final int mainSheet, final Map<String, Integer> mainHeader,
			int mainHeaderRow, File fileName) {

		// Scan the list of xml levels
		int mainRow = mainHeaderRow + 1;

		try {
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

			// Write the XML to a file
			// root elements
			Document doc = docBuilder.newDocument();
			Element rootElement = doc.createElement("Sources");
			doc.appendChild(rootElement);

			boolean isSuccessful = true;
			while (!XLSUtil.isEndRow(wb, mainSheet, mainRow)) {

				Element srcElement = doc.createElement("Source");
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ID, true);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_BILLBOARD,
						true);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_PRIVACY, true);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_NAME, true);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_LINK, false);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_OWNER, false);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_DESCRIPTION, false);
				isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ENABLED, false);

				for (PostFieldType fieldType : PostFieldType.values()) {
					if ((fieldType.getPersistency() == Persistency.Source) &&
							(mainHeader.containsKey(fieldType.name())))
						isSuccessful &= moveNodeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ENABLED, false);
				}
			}

			if (isSuccessful == true) {

				File xmlFilename = new File(fileName.getPath());

				// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer = transformerFactory.newTransformer();
				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new PrintStream(xmlFilename));
				transformer.setOutputProperty(OutputKeys.INDENT, "yes");
				transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
				transformer.transform(source, result);

			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static boolean moveAttributeToXml(Element srcElement, Workbook wb, int mainSheet, int mainRow,
			Map<String, Integer> mainHeader, String columnName, boolean isRequired) {
		String value = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(columnName));
		if (!value.isEmpty()) {
			StringBuilder attribName = new StringBuilder(columnName);
			attribName.setCharAt(0, Character.toLowerCase(attribName.charAt(0)));
			srcElement.setAttribute(attribName.toString(), value);
		} else if (isRequired == true) // && (value.isEmpty())
		{
			Logger.getGlobal().severe(
					"Unable to find attribute " + columnName + " for row " + (mainRow + 1) + ". Writing failed.");
			return false;
		}
		return true;
	}

	private static boolean moveNodeToXml(Element srcElement, Workbook wb, int mainSheet, int mainRow,
			Map<String, Integer> mainHeader, String columnName, boolean isRequired) {
		String value = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(columnName));
		if (!value.isEmpty()) {
			Element node = srcElement.getOwnerDocument().createElement(columnName);
			srcElement.appendChild(node);
			node.setAttribute("value", value);
		} else if (isRequired == true) // && (value.isEmpty())
		{
			Logger.getGlobal().severe(
					"Unable to find property " + columnName + " for row " + (mainRow + 1) + ". Writing failed.");
			return false;
		}
		return true;
	}
}
