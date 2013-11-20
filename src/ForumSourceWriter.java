import java.io.File;
import java.io.PrintStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.Variant;
import com.pairapp.engine.parser.data.VariantEnum;
import com.pairapp.engine.parser.data.VariantType;
import com.pairapp.engine.parser.data.PostFieldType.Persistency;
import com.pairapp.utilities.LogLineFormatter;

import utility.XLSUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;
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
	final static String COLUMN_REMARKS = "Remarks";
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
			Logger.getGlobal().setUseParentHandlers(false);
			// getParent().removeHandler(Logger.getGlobal().getParent().getHandlers()[0]);
			ConsoleHandler handler = new ConsoleHandler();
			handler.setFormatter(new LogLineFormatter());
			Logger.getGlobal().addHandler(handler);
			Logger.getGlobal().setLevel(Level.INFO);
			
			Logger.getGlobal().info("Work Started.");
			boolean isSuccessful = processFile(inputFile);
			Logger.getGlobal().info("Work " + (isSuccessful ? "finished successfully." : "failed."));
		}

	}

	private static boolean processFile(File inputFile) {
		Workbook wb = XLSUtil.openXLS(inputFile);
		Logger.getGlobal().info("XLS file loaded.");
		
		int mainSheet = 0;

		// Find the header row
		Map<String, Integer> mainHeader = new HashMap<String, Integer>();
		int mainHeaderRow = XLSUtil.getHeaderRow(wb, mainSheet, mainHeader, COLUMN_ID, COLUMN_BILLBOARD,
				COLUMN_PRIVACY, COLUMN_NAME, COLUMN_ENABLED);
		
		if (mainHeaderRow == -1) {
			Logger.getGlobal().info("Severe Started.");
			return false;
		} else {
			
			String sourceFilePath = "ForumSources.xml";
			if (inputFile.getParent() != null)
				sourceFilePath = inputFile.getParent() + "\\" + sourceFilePath;
			return exportToXML(wb, mainSheet, mainHeader, mainHeaderRow, new File(sourceFilePath));
		}
	}

	private static boolean exportToXML(final Workbook wb, final int mainSheet, final Map<String, Integer> mainHeader,
			int mainHeaderRow, File fileName) {
		boolean isSuccessful = true;
		// Scan the list of xml levels
		int mainRow = mainHeaderRow + 1;

		try {
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			
			//Check for redundent coluymns
			ArrayList<String> redundentColumns = new ArrayList(mainHeader.keySet());
			redundentColumns.remove(COLUMN_ORDER);
			redundentColumns.remove(COLUMN_REMARKS);
			redundentColumns.remove(COLUMN_ID);
			redundentColumns.remove(COLUMN_BILLBOARD);
			redundentColumns.remove(COLUMN_PRIVACY);
			redundentColumns.remove(COLUMN_NAME);
			redundentColumns.remove(COLUMN_LINK);
			redundentColumns.remove(COLUMN_OWNER);
			redundentColumns.remove(COLUMN_DESCRIPTION);
			redundentColumns.remove(COLUMN_ENABLED);
			for (PostFieldType fieldType : PostFieldType.values())
				if (fieldType.getPersistency() == Persistency.Source)
						redundentColumns.remove(fieldType.name());
			if (redundentColumns.isEmpty() == false)
			{
				Logger.getGlobal().warning("The XLS data contains the following unknown rows " + Arrays.toString(redundentColumns.toArray(new String[redundentColumns.size()])) + ".");
			}
			
			// Write the XML to a file
			// root elements
			Document doc = docBuilder.newDocument();
			Element rootElement = doc.createElement("Sources");
			doc.appendChild(rootElement);

			while (!XLSUtil.isEndRow(wb, mainSheet, mainRow)) {

				String enabled = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_ENABLED));
				if (Boolean.valueOf(enabled) == true)
				{
					Element srcElement = doc.createElement("Source");
					rootElement.appendChild(srcElement);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ID, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_BILLBOARD, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_PRIVACY, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_NAME, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_LINK, false);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_OWNER, false);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_DESCRIPTION, false);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ENABLED, false);
	
					for (PostFieldType fieldType : PostFieldType.values()) {
						if ((fieldType.getPersistency() == Persistency.Source) &&
								(mainHeader.containsKey(fieldType.name())))
							isSuccessful &= moveNodeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, fieldType, false);
					}
				}
				++mainRow;
				
			}

			if (isSuccessful == true) {

				Logger.getGlobal().info("Beginning to write XML file.");
				
				File xmlFilename = new File(fileName.getPath());

				// write the content into xml file
				TransformerFactory transformerFactory = TransformerFactory.newInstance();
				Transformer transformer = transformerFactory.newTransformer();
				DOMSource source = new DOMSource(doc);
				StreamResult result = new StreamResult(new PrintStream(xmlFilename));
				transformer.setOutputProperty(OutputKeys.INDENT, "yes");
				transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
				transformer.transform(source, result);

				Logger.getGlobal().info("XML file written.");
				
			}
		} catch (Exception e) {
			e.printStackTrace();
			isSuccessful = false;
		}
		return isSuccessful;
		

	}

	private static boolean moveAttributeToXml(Element srcElement, Workbook wb, int mainSheet, int mainRow,
			Map<String, Integer> mainHeader, String columnName, boolean isRequired) {
		String value = mainHeader.containsKey(columnName) ? XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(columnName)) : "";
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
			Map<String, Integer> mainHeader, PostFieldType property, boolean isRequired) {
		String value = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(property.name()));
		if (!value.isEmpty()) {
			if (!isValidPropertyValue(property, value))
			{
				Logger.getGlobal().severe(
						"Unable to set property " + property.name() + " with a value of " + value + " for row " + (mainRow + 1) + ". Writing failed.");
				return false;
			}
			else
			{
				Element node = srcElement.getOwnerDocument().createElement(property.name());
				srcElement.appendChild(node);
				node.setAttribute("value", value);
				return true;
			}
		} else if (isRequired == true) // && (value.isEmpty())
		{
			Logger.getGlobal().severe(
					"Unable to find property " + property.name() + " for row " + (mainRow + 1) + ". Writing failed.");
			return false;
		}
		return true;
	}

	private static boolean isValidPropertyValue(PostFieldType property, String value) {
		if (property.getValueType().getEnumClass() != null)
		{
			return VariantEnum.generateInstance(property.getValueType(), value) != null;
		}
		return true;
	}
}
