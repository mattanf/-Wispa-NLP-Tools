package app;
import java.io.File;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Objects;
import com.pairapp.datalayer.SourcesDatalayer;
import com.pairapp.dataobjects.ForumSource;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.PostFieldType.Persistency;
import com.pairapp.engine.parser.data.VariantTypeEnums.Billboard;
import com.pairapp.utilities.LogLineFormatter;

import utility.XLSUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
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
			//datastoreHelper.setUp();
			doWriteXmlFile = Arrays.asList(args).contains("-w");
			boolean isSuccessful = processFile(inputFile);
			datastoreHelper.tearDown();
			Logger.getGlobal().info("Work " + (isSuccessful ? "finished successfully." : "failed."));
		}

	}

	private static boolean processFile(File inputFile) {
		return loadSourcesFromFile(inputFile) != null;
		 
	}

	public static List<ForumSource> loadSourcesFromFile(File inputFile) {
		datastoreHelper.setUp();
		Workbook wb = XLSUtil.openXLS(inputFile);
		Logger.getGlobal().info("XLS file loaded.");
		
		int mainSheet = 0;

		// Find the header row
		Map<String, Integer> mainHeader = new HashMap<String, Integer>();
		int mainHeaderRow = XLSUtil.getHeaderRow(wb, mainSheet, mainHeader, COLUMN_ID, COLUMN_BILLBOARD,
				COLUMN_PRIVACY, COLUMN_NAME, COLUMN_ENABLED);
		
		
		if (mainHeaderRow == -1) {
			Logger.getGlobal().severe("Unable to locate main row in xls.");
			return null;
		}
		
		if (mainHeader.get(COLUMN_PARSER_MESSAGE) == null)
		{
			XLSUtil.addColumnToHeader(mainHeader, COLUMN_PARSER_MESSAGE);
			XLSUtil.updateHeader(wb, mainSheet, mainHeaderRow, mainHeader);
		}
		
		ArrayList<ForumSource> retSources = new ArrayList<ForumSource>();
		// Scan the list of xml levels
		int mainRow = mainHeaderRow + 1;
		boolean hasExcelChanged = false;
		
		//Check for redundant columns
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
			if ((fieldType.getPersistency() == Persistency.Source) || (fieldType.getPersistency() == Persistency.SourceAssociated))
					redundentColumns.remove(fieldType.name());
		if (redundentColumns.isEmpty() == false)
		{
			Logger.getGlobal().warning("The XLS data contains the following unknown rows " + Arrays.toString(redundentColumns.toArray(new String[redundentColumns.size()])) + ".");
		}
			
			
		while (!XLSUtil.isEndRow(wb, mainSheet, mainRow)) {

			boolean isProvisioningEnabled = Boolean.valueOf(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_ENABLED)));
			if (isProvisioningEnabled == true)
			{
				ForumSource forumSource = new ForumSource(Billboard.toEnum(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_BILLBOARD))),
						XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_ID)),
						XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_NAME)),							
						com.pairapp.dataobjects.ForumPrivacy.toEnum(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PRIVACY))));
				forumSource.setIsProvisionEnabled(isProvisioningEnabled);
				
				boolean isValid = true;
				String parserMessage = null;
				for (PostFieldType fieldType : PostFieldType.values()) {
					if (((fieldType.getPersistency() == Persistency.Source) || (fieldType.getPersistency() == Persistency.SourceAssociated)) &&
							(mainHeader.containsKey(fieldType.name()))) {
						String propertyName = fieldType.name();
						String propertyValue = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(fieldType.name()));
						propertyValue = hackPropertyValueToSingleValue(propertyValue,propertyName);
						if (propertyValue.isEmpty() == false)
						{
							boolean isSet = forumSource.setProperty(propertyName, propertyValue);
							if (isSet == false)
							{
								parserMessage = "Unable to set property " + propertyName + " to value " + propertyValue;
								isValid = false;
							}
						}
						//isSuccessful &= moveNodeToXml(srcElement, wb, mainSheet, mainRow, mainHeader, fieldType, false);
					}
				}
				StringBuilder message = new StringBuilder();
				if (isValid == true)
				{
					isValid = SourcesDatalayer.validateSource(message, forumSource, false);
					parserMessage = isValid ? "OK" : message.toString();
				}
				if (Objects.equals(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE)), parserMessage) == false)
				{
					XLSUtil.setCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE), parserMessage);
					hasExcelChanged = true;
				}
				
				//Add the information of the source to the xml
				if (isValid)
				{
					retSources.add(forumSource);
				}
				else {
					Logger.getGlobal().severe("Source " + forumSource + ": " + parserMessage);
				}
			}
			else
			{
				if (Objects.equals(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE)), "Disabled") == false)
				{
					XLSUtil.setCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE), "Disabled");
					hasExcelChanged = true;
				}
			}
			++mainRow;
			
		}

		if (hasExcelChanged)
			XLSUtil.saveXLS(wb, inputFile);
		datastoreHelper.tearDown();
		return retSources;
		

	}

	private static String hackPropertyValueToSingleValue(String propertyValue, String propertyName) {
		if ((PostFieldType.ForumLocationState.name().equals(propertyName)) ||
				(PostFieldType.ForumLocationRegion.name().equals(propertyName)) ||
				(PostFieldType.ForumLocationSubRegion.name().equals(propertyName)) ||
				(PostFieldType.ForumLocationCity.name().equals(propertyName))) {
			int delim = propertyValue.indexOf(';');
			if (delim != -1)
				return propertyValue.substring(0, delim);
		}
		return propertyValue;
	}
}
