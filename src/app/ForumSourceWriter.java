package app;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Workbook;

import utility.XLSUtil;

import com.google.appengine.api.xmpp.MessageBuilder;
import com.google.appengine.tools.development.testing.LocalDatastoreServiceTestConfig;
import com.google.appengine.tools.development.testing.LocalServiceTestHelper;
import com.pairapp.datalayer.SourcesDatalayer;
import com.pairapp.dataobjects.ForumSource;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.PostFieldType.Persistency;
import com.pairapp.engine.parser.data.PostFields;
import com.pairapp.engine.parser.data.VariantTypeEnums.Billboard;
import com.pairapp.engine.parser.data.VariantTypeEnums.CountryState;
import com.pairapp.utilities.LogLineFormatter;
import com.pairapp.utilities.StringUtilities;
import com.pairapp.utilities.Utils;
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
			Logger.getGlobal().info("Work " + (isSuccessful ? "finished successfully." : "failed."));	
			//datastoreHelper.tearDown();
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
				StringBuilder message = new StringBuilder();
				for (PostFieldType fieldType : PostFieldType.values()) {
					if (((fieldType.getPersistency() == Persistency.Source) || (fieldType.getPersistency() == Persistency.SourceAssociated)) &&
							(mainHeader.containsKey(fieldType.name()))) {
						String propertyName = fieldType.name();
						String propertyValue = XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(fieldType.name()));
						propertyValue = handleMultiStateProperty(forumSource,propertyValue,propertyName,message);
						//Handle error
						if (propertyValue == null)
						{
							isValid = false;
						}
						else if (propertyValue.isEmpty() == false)
						{
							boolean isSet = forumSource.setProperty(propertyName, propertyValue);
							if (isSet == false)
							{
								if (message.length() != 0) message.append("\n");
								message.append("Unable to set property " + propertyName + " to value " + propertyValue);
								isValid = false;
							}
						}
					}
				}
				
				if (isValid == true)
				{
					PostFields newFields = SourcesDatalayer.extrapolateMissingSourceFields(forumSource, message, false);
					if (newFields != null)
					{
						forumSource.setProperties(newFields);
					}
					else {
						isValid = false;
					}
				}
				
				if (isValid == true)
				{
					isValid = SourcesDatalayer.validateSource(message, forumSource, false);
				}
				
				String excelMessage = message.toString();
				if (isValid == true)
					excelMessage = "OK";
				else if (message.length() == 0)
					excelMessage = "Unknown Error";
					
				if (Objects.equals(XLSUtil.getCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE)), excelMessage) == false)
				{
					XLSUtil.setCellString(wb, mainSheet, mainRow, mainHeader.get(COLUMN_PARSER_MESSAGE), excelMessage);
					hasExcelChanged = true;
				}
				
				//Add the information of the source to the xml
				if (isValid)
				{
					retSources.add(forumSource);
				}
				else {
					Logger.getGlobal().severe("Source " + forumSource + ": " + excelMessage);
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

	private static String handleMultiStateProperty(ForumSource source, String propertyValue, String propertyName, StringBuilder message) {
		String retValue = propertyValue;
		if (PostFieldType.ForumLocationState.name().equals(propertyName)) {
			String splittedProperties[] = propertyValue.split("\\s*;\\s*");
			if (splittedProperties.length > 1)
			{
				//If we do not know which state it is return no state
				retValue = "";
				//We have multiple values for states
				String finalPropString = "";
				for(String indProp : splittedProperties)
				{
					CountryState countryState = CountryState.freeTextToEnum(indProp);
					//hack to report error
					if (countryState == null)
					{
						if (message.length() != 0) message.append("\n");
						message.append("Unknown country " + indProp + " in ForumLocationState");
						retValue = null;
					}
					else finalPropString = StringUtilities.concate(finalPropString,  ";", countryState.name());
				}
				if (finalPropString.isEmpty() == false)
					source.setProperty(PostFieldType.ANACountryStates.name(), finalPropString);
			}
		}
		return retValue;
	}
}
