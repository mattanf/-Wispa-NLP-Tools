import java.io.File;
import java.io.PrintStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.google.appengine.labs.repackaged.com.google.common.base.Objects;
import com.pairapp.datalayer.SourcesDatalayer;
import com.pairapp.dataobjects.ForumPrivacy;
import com.pairapp.dataobjects.ForumSource;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.VariantEnum;
import com.pairapp.engine.parser.data.PostFieldType.Persistency;
import com.pairapp.engine.parser.data.VariantTypeEnums.Billboard;
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
import com.google.appengine.tools.development.testing.LocalDatastoreServiceTestConfig;
import com.google.appengine.tools.development.testing.LocalServiceTestHelper;
public class ForumSourceWriter {

	final static String COLUMN_ORDER = "Order";
	final static String COLUMN_REMARKS = "Remarks";
	final static String COLUMN_ADDED_ON = "Added On";
	final static String COLUMN_ID = "Id";
	final static String COLUMN_BILLBOARD = "Billboard";
	final static String COLUMN_PRIVACY = "Privacy";
	final static String COLUMN_NAME = "Name";
	final static String COLUMN_LINK = "Link";
	final static String COLUMN_OWNER = "Owner";
	final static String COLUMN_DESCRIPTION = "Description";
	final static String COLUMN_ENABLED = "Enabled";
	final static String COLUMN_PARSER_MESSAGE = "Parser Message";
	static boolean doWriteXmlFile = false;
	
	private static final LocalServiceTestHelper datastoreHelper = new LocalServiceTestHelper(
			new LocalDatastoreServiceTestConfig());
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		File inputFile = new File(args.length == 0 ? "" : args[args.length - 1]);
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
			datastoreHelper.setUp();
			doWriteXmlFile = Arrays.asList(args).contains("-w");
			boolean isSuccessful = processFile(inputFile);
			datastoreHelper.tearDown();
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
			
			if (mainHeader.get(COLUMN_PARSER_MESSAGE) == null)
			{
				XLSUtil.addColumnToHeader(mainHeader, COLUMN_PARSER_MESSAGE);
				XLSUtil.updateHeader(wb, mainSheet, mainHeaderRow, mainHeader);
			}
			
			String sourceFilePath = "ReferenceData\\ForumSources.xml";
			//if (inputFile.getParent() != null)
			//	sourceFilePath = inputFile.getParent() + "\\" + sourceFilePath;
			return exportToXML(wb, mainSheet, mainHeader, mainHeaderRow, new File(sourceFilePath), inputFile);
		}
	}

	private static boolean exportToXML(final Workbook wb, final int mainSheet, final Map<String, Integer> mainHeader,
			int mainHeaderRow, File xmlFileName, File xlsFile) {
		boolean isSuccessful = true;
		// Scan the list of xml levels
		int mainRow = mainHeaderRow + 1;
		boolean hasExcelChanged = false;
		try {
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
			
			//Check for redundent coluymns
			ArrayList<String> redundentColumns = new ArrayList<String>(mainHeader.keySet());
			redundentColumns.remove(COLUMN_ORDER);
			redundentColumns.remove(COLUMN_REMARKS);
			redundentColumns.remove(COLUMN_ADDED_ON);
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
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ID, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_BILLBOARD, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_PRIVACY, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_NAME, true);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_LINK, false);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_OWNER, false);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_DESCRIPTION, false);
					isSuccessful &= moveAttributeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, COLUMN_ENABLED, false);
					
					ForumSource forumSource = new ForumSource(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_ID)),
							XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_NAME)),
							Billboard.toEnum(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_BILLBOARD))),
							XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_LINK)),
							XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_OWNER)),
							ForumPrivacy.toEnum(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PRIVACY))));
					
					for (PostFieldType fieldType : PostFieldType.values()) {
						if ((fieldType.getPersistency() == Persistency.Source) &&
								(mainHeader.containsKey(fieldType.name()))) {
							if (XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(fieldType.name())).isEmpty() == false)
								forumSource.setProperty(fieldType.name(), XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(fieldType.name())));
							isSuccessful &= moveNodeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, fieldType, false);
						}
					}
					StringBuilder message = new StringBuilder();
					boolean isValid = SourcesDatalayer.validateSource(message, forumSource, false);
					String parserMessage = isValid ? "OK" : message.toString();
					if (Objects.equal(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE)), parserMessage) == false)
					{
						XLSUtil.setCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE), parserMessage);
						hasExcelChanged = true;
					}
					
					//Add the information of the source to the xml
					if (isValid)
					{
						rootElement.appendChild(srcElement);
					}
				}
				else
				{
					if (Objects.equal(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE)), "Disabled") == false)
					{
						XLSUtil.setCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE), "Disabled");
						hasExcelChanged = true;
					}
				}
				++mainRow;
				
			}

			if (hasExcelChanged)
				XLSUtil.saveXLS(wb, xlsFile);
			
			if (isSuccessful == true && doWriteXmlFile) {

				Logger.getGlobal().info("Beginning to write XML file.");
				
				File xmlFilename = new File(xmlFileName.getPath());

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
		value = value.trim();
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
		value = value.trim();
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
