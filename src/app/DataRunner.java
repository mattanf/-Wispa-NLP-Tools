package app;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;

import java.util.logging.Logger;


import com.google.appengine.tools.development.testing.LocalDatastoreServiceTestConfig;
import com.google.appengine.tools.development.testing.LocalServiceTestHelper;
import com.google.appengine.tools.remoteapi.RemoteApiInstaller;
import com.google.appengine.tools.remoteapi.RemoteApiOptions;
import com.google.apphosting.api.ApiProxy;
import com.google.apphosting.api.ApiProxy.Environment;
import com.pairapp.dataobjects.Comment;
import com.pairapp.engine.parser.MessageParser;
import com.pairapp.engine.parser.data.PostData;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.Variant;
import com.pairapp.engine.parser.data.VariantDate;
import com.pairapp.engine.parser.data.VariantEnum;
import com.pairapp.engine.parser.data.VariantType;
import com.pairapp.engine.parser.data.VariantTypeEnums;
import com.pairapp.engine.parser.location.GeocodeQuerier;
import com.pairapp.utilities.LogLineFormatter;


public class DataRunner {

	static class TestEnvironment implements ApiProxy.Environment  {
		  @Override
		  public String getAppId() {
		    return "test";
		  }
		  @Override
		  public String getVersionId() {
		    return "1.0";
		  }
		  @Override
		  public String getRequestNamespace() {
		    return "";
		  }
		  @Override
		  public String getAuthDomain() {
		    throw new UnsupportedOperationException();
		  }
		  @Override
		  public boolean isLoggedIn() {
		    throw new UnsupportedOperationException();
		  }
		  @Override
		  public String getEmail() {
		    throw new UnsupportedOperationException();
		  }
		  @Override
		  public boolean isAdmin() {
		    throw new UnsupportedOperationException();
		  }
		  @Override
		  public Map<String, Object> getAttributes() {
		    return new HashMap<String, Object>();
		  }
		@Override
		public long getRemainingMillis() {
			return 111110;
		}
		//@Override
		public String getModuleId() {
			return null;
		}
		}
	
	final static int STYLE_CENTER = 1 << 0;
	final static int STYLE_BORDER_RIGHT = 1 << 1;
	final static int STYLE_BORDER_LEFT = 1 << 2;
	final static int STYLE_BORDER_BOTTOM = 1 << 3;
	final static int STYLE_PERCENTAGE = 1 << 4;
	
	final static int STYLE_COLOR_LIGHT = 1 << 16;
	final static int STYLE_COLOR_GREEN = 1 << 17;
	final static int STYLE_COLOR_YELLOW = 1 << 18;
	final static int STYLE_COLOR_RED = 1 << 19;
	final static int STYLE_COLOR_GRAY = 1 << 20;
	final static int STYLE_COLOR = STYLE_COLOR_GREEN | STYLE_COLOR_YELLOW | STYLE_COLOR_RED | STYLE_COLOR_GRAY;
	
	private final static String SUBJECT_MESSAGE = "Message";
	private final static int MAX_HEADERS = 100;

	private static PostData basePostData;
	
	static ApiProxy.Environment testEnvironment = new TestEnvironment();
	private static final LocalServiceTestHelper datastoreHelper = new LocalServiceTestHelper(
			new LocalDatastoreServiceTestConfig());

	private static int startParseRow = -1;
	private static int endParseRow = 10000000;

	private static HashMap<Integer, CellStyle> cellStyles;
	private static int usedColorIndexRunner = 0;

	static class ParseThreadData
	{
		private int row;
		private PostData inpData;
		private PostData outData;
		
		public ParseThreadData(int row, PostData inpData, PostData outData) {
			this.row = row;
			this.inpData = inpData;
			this.outData = outData;
				
		}
		public int getRow() {
			return row;
		}
		public PostData getInput() {
			return inpData;
		}
		public PostData getOutput() {
			return outData;
		}
	}
	
	static final int NUMBER_OF_THREADS = Runtime.getRuntime().availableProcessors();
	static ArrayBlockingQueue<ParseThreadData> threadInputQueue = new ArrayBlockingQueue<>(NUMBER_OF_THREADS + 1);
	static ArrayBlockingQueue<ParseThreadData> threadOutputQueue = new ArrayBlockingQueue<>((NUMBER_OF_THREADS + 1) * 2);
	static ThreadLocal<MessageParser> threadMessageParser = new ThreadLocal<MessageParser>();
	static ExecutorService executorService = Executors.newFixedThreadPool(NUMBER_OF_THREADS);
	static boolean workThreaded = true;
	static boolean connectToLocalServer = false;
	
	// private static int endParseRow = -1;
	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		
		/*foo(4);
		foo(3);
		foo(2);
		int r = 2;
		while (r < 6)
		{
			
			Thread.sleep(1000);
			
		}*/
		//foo();
		// This class does in a round about way initializes the log. we need to call it before making changes to the
		// loging mechnism
		// GeocodeQuerier.instance();
		// Remove the previous logging mechanism and add a better handler
		Logger.getGlobal().setUseParentHandlers(false);
		// getParent().removeHandler(Logger.getGlobal().getParent().getHandlers()[0]);
		ConsoleHandler handler = new ConsoleHandler();
		handler.setFormatter(new LogLineFormatter());
		Logger.getGlobal().addHandler(handler);
		Logger.getGlobal().setLevel(Level.WARNING);

		int fileIndex = 0;
		for (int i = 0; i < args.length; ++i) {
			if (args[fileIndex].startsWith("-")) {
				if (args[fileIndex].matches("-r\\d+"))
					endParseRow = startParseRow = Integer.valueOf(args[fileIndex].substring(2)) - 1;
				if (args[fileIndex].matches("-rs\\d+"))
					startParseRow = Integer.valueOf(args[fileIndex].substring(3)) - 1;
				if (args[fileIndex].matches("-re\\d+"))
					endParseRow = Integer.valueOf(args[fileIndex].substring(3)) - 1;
				if (args[fileIndex].matches("-local"))
					connectToLocalServer = true;
				++fileIndex;
			} else
				break;
		}

		if (args.length - fileIndex == 0) {
			System.out.println("Usage: [XLS/X source file]");
		} else {
			fileIndex = args.length - 1;
			File sourceFile = new File(args[fileIndex]);

			if ((args[fileIndex].toLowerCase().endsWith(".xls") == false) &&
					(args[fileIndex].toLowerCase().endsWith(".xlsx") == false))
				System.err.println("Error: Specified input file '" + args[fileIndex] + "' is not a xls/x file.");
			else if (!sourceFile.isFile() || !sourceFile.exists())
				System.err.println("Error: Specified file '" + args[fileIndex] + "' could not be read.");
			else {
				String secondaryFileName = args[fileIndex];
				if (connectToLocalServer)
					connectToLocalServer = connectToLocalServer();
				GeocodeQuerier.instance().setEnableGoogleQuerying(connectToLocalServer);
				secondaryFileName = extendFileName(secondaryFileName, "-compiled");
				File targetFile = new File(secondaryFileName);
				
				cellStyles = new HashMap<>();
				System.out.println("Started parse process");
				datastoreHelper.setUp();
				Environment prevEnv = ApiProxy.getCurrentEnvironment();
				ApiProxy.setEnvironmentForCurrentThread(testEnvironment);

				analyzeFile(sourceFile, targetFile);
				ApiProxy.setEnvironmentForCurrentThread(prevEnv);
				datastoreHelper.tearDown();
				System.out.println("Ended parse process");
			}
		}
	}

	
	
    public static boolean connectToLocalServer() throws IOException  {
        try {
	    	String username = "mattanf@gmail.com";//System.console().readLine("username: ");
	        String password = "some_pass";
	         //   new String(System.console().readPassword("password: "));
	        RemoteApiOptions options = new RemoteApiOptions()
	            //.server("wispa-test.appspot.com", 443)
	        	.server("localhost", 8888)
	            .credentials(username, password);
	        RemoteApiInstaller installer = new RemoteApiInstaller();
	        installer.install(options);
	        options.reuseCredentials(username, installer.serializeCredentials());
        }
        catch (Throwable e)
        {
        	Logger.getGlobal().severe("Could not connect to server. Location will not be extracted fully");
        	return false;
        }
        return true;
    }
	private static String extendFileName(String fileName, String extention) {
		int delimPer = fileName.lastIndexOf('.');
		if (delimPer != -1)
			return fileName.substring(0, delimPer) + extention + fileName.substring(delimPer);
		else
			return fileName + extention;
	}

	private static boolean analyzeFile(File source, File target) throws Exception {
		boolean isSuccess = true;
		try {

			Workbook wb = openXLS(source);
			MessageParser parser = new MessageParser();
			int readRowNum = 0;

			/*
			 * Read the pre fields header
			 */
			basePostData = new PostData();
			while (!getCellString(wb, 0, readRowNum, 0).equalsIgnoreCase(SUBJECT_MESSAGE) && isSuccess) {
				String h1 = getCellString(wb, 0, readRowNum, 0).trim();
				String h2 = getCellString(wb, 0, readRowNum, 1);

				if ((!h1.isEmpty()) && (!h2.isEmpty()) && (!h1.contains(" "))) {
					basePostData.addField(h1, null, h2);
				}
				++readRowNum;
				isSuccess = !isEndRow(wb, 0, readRowNum);
			}
			if (!isSuccess)
				System.err.println("Error: ParserTree could not be find header message.");
			else {
				isSuccess = parser.init(connectToLocalServer);
				if (!isSuccess)
					System.err.println("Error: ParserTree could not be initialized. Data source could not be found.");
				else {
					HashMap<String, Integer> metaDataHeaders = new HashMap<>();
					HashMap<String, Integer> outputHeaders = new HashMap<>();

					// Read the header
					int sectionInt = 0;
					int rowWidth = getRowWidth(wb, 0, readRowNum);
					int headerRowNum = readRowNum;
					for (int i = 0; i < rowWidth; ++i) {
						String columnName = getCellString(wb, 0, readRowNum, i).trim();
						if (columnName.matches("Column\\d+"))
							columnName = "";
						switch (sectionInt) {
						case 0:
							if ((columnName.isEmpty() == false) &&
									(columnName.compareToIgnoreCase(SUBJECT_MESSAGE) != 0))
								metaDataHeaders.put(columnName, i);
							if (columnName.isEmpty() == true)
								sectionInt = 1;
							break;
						case 1:
							outputHeaders.put(columnName, i);
							break;
						}
					}
					++readRowNum;

					int countMessageDiffer[] = new int[MAX_HEADERS];
					Arrays.fill(countMessageDiffer, 0);

					int firstMessageRow = Math.max(readRowNum, startParseRow);
					TreeMap<String, HashMap<String, Integer>> resultCount = new TreeMap<>(String.CASE_INSENSITIVE_ORDER);
					
					for (int stage = 0; stage < 2; ++stage) {
						readRowNum = firstMessageRow;

						// Compare and write the report
						while (!isEndRow(wb, 0, readRowNum)) {
							if (readRowNum > endParseRow)
								break;

							if (((readRowNum - firstMessageRow) % 100) == 0)
							{	
								Level prevLevel = Logger.getGlobal().getLevel();
								Logger.getGlobal().setLevel(Level.INFO);
								Logger.getGlobal().info(
										"Parsing row " + (readRowNum - firstMessageRow) + " in stage " + (stage + 1));
								Logger.getGlobal().setLevel(prevLevel);
							}

							if (stage == 0)
								parseRowMessage(wb, headerRowNum, readRowNum, metaDataHeaders, outputHeaders, parser, isEndRow(wb, 0, readRowNum + 1) || (readRowNum == endParseRow));
							else
								compareRowRecord(wb, readRowNum, metaDataHeaders, outputHeaders, resultCount);

							++readRowNum;

						}
					}
					saveComparisonResults(wb, resultCount);
					saveWorkbook(wb, target);
				}
			}
		} catch (IOException e) {
			System.err.println("Error: " + e.getMessage());

			isSuccess = false;
		}

		return isSuccess;
	}

	private static void compareRowRecord(Workbook wb, int readRowNum, HashMap<String, Integer> metaDataHeaders,
			HashMap<String, Integer> outputHeaders, TreeMap<String, HashMap<String, Integer>> resultCount) {
		String messageStr = getCellString(wb, 0, readRowNum, 0);
		if (messageStr.isEmpty() == false) {
			
			String isValidString = metaDataHeaders.containsKey("PostIsValid") ? getCellString(wb, 0, readRowNum, metaDataHeaders.get("PostIsValid")) : "True";
			boolean isValid = Boolean.valueOf(isValidString);
			String purposeString =  metaDataHeaders.containsKey("PostPurpose") ? getCellString(wb, 0, readRowNum,
					metaDataHeaders.get("PostPurpose")) : "Unknown";
			boolean isOffer = purposeString.equalsIgnoreCase("Offer");
			boolean isSeek = purposeString.equalsIgnoreCase("Seek");
			boolean isSwap = purposeString.equalsIgnoreCase("Swap");
			
			
			Iterator<Entry<String, Integer>> outputHeaderIt = outputHeaders.entrySet().iterator();
			while (outputHeaderIt.hasNext()) {
				Entry<String, Integer> outputHeader = outputHeaderIt.next();
				String fieldName = outputHeader.getKey();
				Integer origHeaderPos = metaDataHeaders.get(fieldName);
				boolean isValidAgnosticField = isValidAgnosticFieldName(fieldName);
				boolean isBooleanColumn = (PostFieldType.toEnum(fieldName) != null) && (PostFieldType.toEnum(fieldName).getValueType().equals(VariantType.Boolean));
						
				if (origHeaderPos != null) {
					
					// Compare the values
					String genValue = normNumeric(getCellString(wb, 0, readRowNum, outputHeader.getValue()));
					String origValuesString = getCellString(wb, 0, readRowNum, origHeaderPos);
					String origValues[] = origValuesString.split("\\|");
					int wrongness = 3;
					
					for(String origValue : origValues)
					{	
						if ((!origValue.isEmpty()) || (origValues.length == 1))
						{
							if (origValue.equalsIgnoreCase("[Empty]"))
								origValue = "";
							
							origValue = normNumeric(origValue);
							boolean isFalseNegative = (origValue.isEmpty() == false) && (genValue.isEmpty() == true);
							boolean isFalsePositive = !isFalseNegative && origValue.compareToIgnoreCase(genValue) != 0;
							if (isBooleanColumn || (isBooleanString(origValue) && isBooleanString(genValue)))
							{
								isFalseNegative = (Boolean.valueOf(origValue) == true) && (Boolean.valueOf(genValue) == false);
								isFalsePositive = (Boolean.valueOf(origValue) == false) && (Boolean.valueOf(genValue) == true);
							}
							
							boolean isSame = !isFalseNegative && !isFalsePositive;
							
							
							if (isFalsePositive) 
								wrongness = Math.min(wrongness, 2);
							else if (isFalseNegative) 
								wrongness = Math.min(wrongness, 1);
							else if (isSame)
								wrongness = Math.min(wrongness, 0);
						}
						else
						{
							Logger.getGlobal().severe("Precalculated value of " + outputHeader.getKey() + " in row " + (readRowNum + 1) + " has an illegal pipe symbol. Consider adding the word \"[Empty]\" or removing the pipe");
						}
						
					}
					if (wrongness != 3)
					{
						addComparisonData(fieldName, isValid | isValidAgnosticField, isOffer, isSeek, isSwap, wrongness == 2, wrongness == 1, wrongness == 0,
								resultCount);
						setCellStyle(wb, 0, readRowNum, outputHeader.getValue(), (wrongness >= 2 ? STYLE_COLOR_RED :
								(wrongness >= 1 ? STYLE_COLOR_YELLOW : STYLE_COLOR_GREEN)) | (isValid ? 0 : STYLE_COLOR_LIGHT));
					}
					else
					{
						setCellStyle(wb, 0, readRowNum, outputHeader.getValue(), STYLE_COLOR_GRAY);
					}
							
				}					
			}
		}
	}

	private static boolean isBooleanString(String value) {
		return "False".equalsIgnoreCase(value) || "True".equalsIgnoreCase(value);
	}


	private static String normNumeric(String cellString) {
		if (cellString != null)
		{
			cellString = cellString.trim();
			if (cellString.matches("-?\\d+(,\\d{3})*(\\.0*)?"))
			{
				int delim = cellString.indexOf(".");
				if (delim != -1)
					cellString = cellString.substring(0, delim);
				return Long.toString(Long.parseLong(cellString));
			}
			if (cellString != null && cellString.matches("-?\\d+(,\\d{3})*\\.\\d+"))
			{
				return Double.toString(Double.parseDouble(cellString));
			}
		}
		return cellString.trim();
	}

	private static void addComparisonData(String fieldName, boolean isValid, boolean isOffer, boolean isSeek, boolean isSwap,
			boolean isFalsePositive, boolean isFalseNegative, boolean isSame,
			TreeMap<String, HashMap<String, Integer>> resultCount) {
		HashMap<String, Integer> fieldMap = resultCount.get(fieldName);
		if (fieldMap == null) {
			fieldMap = new HashMap<>();
			resultCount.put(fieldName, fieldMap);
		}
		String baseKey = isFalseNegative ? "FN" : (isFalsePositive ? "FP" : "EQ");
		incrementFieldInMap(fieldMap, baseKey);
		//boolean isValidAgnosticField = isValidAgnosticFieldName(fieldName);
		
		if (isValid == true) {
			incrementFieldInMap(fieldMap, "Valid-" + baseKey);
			if (isSeek == true)
				incrementFieldInMap(fieldMap, "Seek-Valid-" + baseKey);
			if (isOffer == true)
				incrementFieldInMap(fieldMap, "Offer-Valid-" + baseKey);
			if (isSwap == true)
				incrementFieldInMap(fieldMap, "isSwap-Valid-" + baseKey);
		}
	}

	/**
	 * @param fieldMap
	 * @param key
	 */
	private static void incrementFieldInMap(HashMap<String, Integer> fieldMap, String key) {
		Integer value = fieldMap.get(key);
		if (value == null)
			value = new Integer(1);
		else
			value = value + 1;
		fieldMap.put(key, value);
	}

	private static void saveComparisonResults(Workbook wb, TreeMap<String, HashMap<String, Integer>> resultCount) {
		Sheet oldSheet = wb.getSheet("Comparison");
		if (oldSheet != null)
			wb.removeSheetAt(wb.getSheetIndex(oldSheet));
		Sheet compSheet = wb.createSheet("Comparison");
		wb.setSheetOrder("Comparison", 1);
		int sheetNum = wb.getSheetIndex(compSheet);
		
		compSheet.setColumnWidth(0, 4500);
		
		
		// Write the header
		int columnIndex = 0;
		setCellStyle(wb, sheetNum, 1, columnIndex, STYLE_CENTER | STYLE_BORDER_BOTTOM | STYLE_BORDER_RIGHT);
		for (int i = 0; i < 4; ++i) {
			getCell(wb, sheetNum, 1, ++columnIndex, true).setCellValue("FN");
			setCellStyle(wb, sheetNum, 1, columnIndex, STYLE_CENTER | STYLE_BORDER_BOTTOM | STYLE_BORDER_LEFT);
			getCell(wb, sheetNum, 1, ++columnIndex, true).setCellValue("FP");
			setCellStyle(wb, sheetNum, 1, columnIndex, STYLE_CENTER | STYLE_BORDER_BOTTOM);
			getCell(wb, sheetNum, 1, ++columnIndex, true).setCellValue("Equal");
			setCellStyle(wb, sheetNum, 1, columnIndex, STYLE_CENTER | STYLE_BORDER_BOTTOM | STYLE_BORDER_RIGHT);
		}

		setCellStyle(wb, sheetNum, 0, 1, STYLE_CENTER | STYLE_BORDER_RIGHT | STYLE_BORDER_LEFT);
		getCell(wb, sheetNum, 0, 1, true).setCellValue("All");
		setCellStyle(wb, sheetNum, 0, 4, STYLE_CENTER | STYLE_BORDER_RIGHT | STYLE_BORDER_LEFT);
		getCell(wb, sheetNum, 0, 4, true).setCellValue("Only Valid");
		setCellStyle(wb, sheetNum, 0, 7, STYLE_CENTER | STYLE_BORDER_RIGHT | STYLE_BORDER_LEFT);
		getCell(wb, sheetNum, 0, 7, true).setCellValue("Seek (Only Valid)");
		setCellStyle(wb, sheetNum, 0, 10, STYLE_CENTER | STYLE_BORDER_RIGHT | STYLE_BORDER_LEFT);
		setCellStyle(wb, sheetNum, 0, 12, STYLE_CENTER | STYLE_BORDER_RIGHT | STYLE_BORDER_LEFT);
		getCell(wb, sheetNum, 0, 10, true).setCellValue("Offer (Only Valid)");
		compSheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 3));
		compSheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 6));
		compSheet.addMergedRegion(new CellRangeAddress(0, 0, 7, 9));
		compSheet.addMergedRegion(new CellRangeAddress(0, 0, 10, 12));

		// Write the data
		int writeRow = 2;
		Iterator<Entry<String, HashMap<String, Integer>>> fieldIt = resultCount.entrySet().iterator();
		while (fieldIt.hasNext()) {
			Entry<String, HashMap<String, Integer>> compData = fieldIt.next();
			String fieldName = compData.getKey();

			Integer falseNegative = nullToZero(compData.getValue().get("FN"));
			Integer falsePositive = nullToZero(compData.getValue().get("FP"));
			Integer equal = nullToZero(compData.getValue().get("EQ"));
			Integer total = falseNegative + falsePositive + equal;

			Integer falseNegativeValidMessage = nullToZero(compData.getValue().get("Valid-FN"));
			Integer falsePositiveValidMessage = nullToZero(compData.getValue().get("Valid-FP"));
			Integer equalValidMessage = nullToZero(compData.getValue().get("Valid-EQ"));
			Integer totalValidMessage = falseNegativeValidMessage + falsePositiveValidMessage + equalValidMessage;

			Integer falseNegativeValidMessageSeek = nullToZero(compData.getValue().get("Seek-Valid-FN"));
			Integer falsePositiveValidMessageSeek = nullToZero(compData.getValue().get("Seek-Valid-FP"));
			Integer equalValidMessageSeek = nullToZero(compData.getValue().get("Seek-Valid-EQ"));
			Integer totalValidMessageSeek = falseNegativeValidMessageSeek + falsePositiveValidMessageSeek + equalValidMessageSeek;

			Integer falseNegativeValidMessageOffer = nullToZero(compData.getValue().get("Offer-Valid-FN"));
			Integer falsePositiveValidMessageOffer = nullToZero(compData.getValue().get("Offer-Valid-FP"));
			Integer equalValidMessageOffer = nullToZero(compData.getValue().get("Offer-Valid-EQ"));
			Integer totalValidMessageOffer = falseNegativeValidMessageOffer + falsePositiveValidMessageOffer + equalValidMessageOffer;

			columnIndex = -1;
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellValue(fieldName);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(falseNegative + "/" + total);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(falsePositive + "/" + total);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(equal + "/" + total);

			boolean isValidAgnosticField = isValidAgnosticFieldName(fieldName);
			if (isValidAgnosticField == false)
			{
				getCell(wb, sheetNum, writeRow, ++columnIndex, true)
						.setCellFormula(falseNegativeValidMessage + "/" + totalValidMessage);
				getCell(wb, sheetNum, writeRow, ++columnIndex, true)
						.setCellFormula(falsePositiveValidMessage + "/" + totalValidMessage);
				getCell(wb, sheetNum, writeRow, ++columnIndex, true)
						.setCellFormula(equalValidMessage + "/" + totalValidMessage);
			}
			else columnIndex = columnIndex + 3;
			
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(
					falseNegativeValidMessageSeek + "/" + totalValidMessageSeek);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(
					falsePositiveValidMessageSeek + "/" + totalValidMessageSeek);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(
					equalValidMessageSeek + "/" + totalValidMessageSeek);

			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(
					falseNegativeValidMessageOffer + "/" + totalValidMessageOffer);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(
					falsePositiveValidMessageOffer + "/" + totalValidMessageOffer);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(
					equalValidMessageOffer + "/" + totalValidMessageOffer);
			
			for (int i = 0; i < 13; ++i)
				setCellStyle(wb, sheetNum, writeRow, i, STYLE_PERCENTAGE | (i % 3 == 0 ? STYLE_BORDER_RIGHT : 0) | (i % 3 == 1 ? STYLE_BORDER_LEFT : 0));
			
			++writeRow;
		}

	}

	/**
	 * @param fieldName
	 * @return
	 */
	private static boolean isValidAgnosticFieldName(String fieldName) {
		boolean isValidAgnosticField = (fieldName.equalsIgnoreCase("PostIsValid") || fieldName.equalsIgnoreCase("PostIsSpam"));
		return isValidAgnosticField;
	}

	private static void setCellStyle(Workbook wb, int sheetNum, int rowNum, int columnNum, int style) {
		Cell cell = getCell(wb, sheetNum, rowNum, columnNum, true);
		if (cell != null)
		{
			CellStyle cellStyle = cellStyles.get(style);
			if (cellStyle == null)
			{
				cellStyle = wb.createCellStyle();
				if ((style & STYLE_CENTER) != 0)
					cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
				if ((style & STYLE_BORDER_BOTTOM) != 0)
					cellStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
				if ((style & STYLE_BORDER_LEFT) != 0)
					cellStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
				if ((style & STYLE_BORDER_RIGHT) != 0)
					cellStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
				if ((style & STYLE_PERCENTAGE) != 0)
					cellStyle.setDataFormat(wb.createDataFormat().getFormat("0.0%"));
				if ((style & STYLE_COLOR) != 0)
				{
					//Calc the color
					short[] colorRGB;
					if ((style & STYLE_COLOR_GREEN) != 0)
						colorRGB = new short[] {64, 255, 64};
					else if ((style & STYLE_COLOR_YELLOW) != 0)
						colorRGB = new short[] {255, 255, 64};
					else if ((style & STYLE_COLOR_RED) != 0)
						colorRGB = new short[] {255, 64, 64};
					else //if ((style & STYLE_COLOR_GRAY) != 0)
						colorRGB = new short[] {200, 200, 200};
					
					
					if ((style & STYLE_COLOR_LIGHT) != 0)
					{
						colorRGB[0] = (short) (255 - ((255 - colorRGB[0]) / 2));
						colorRGB[1] = (short) (255 - ((255 - colorRGB[1]) / 2));
						colorRGB[2] = (short) (255 - ((255 - colorRGB[2]) / 2));
					}
					
					cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
					if (wb instanceof HSSFWorkbook) {
						HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette();
						// get the color which most closely matches the color you want to use
						
						
						HSSFColor color =  palette.findSimilarColor(colorRGB[0], colorRGB[1], colorRGB[2]);
						if ((color.getTriplet()[0] != colorRGB[0]) || (color.getTriplet()[1] != colorRGB[1]) || (color.getTriplet()[2] != colorRGB[2]))
						{
							try {
								short colorIndex = (short)(63 - usedColorIndexRunner);
								palette.setColorAtIndex((short) colorIndex, (byte)colorRGB[0], (byte)colorRGB[1], (byte)colorRGB[2]);
								HSSFColor tmpColor =  palette.getColor(colorIndex);
								++usedColorIndexRunner;
								
								if (tmpColor != null)
									color = tmpColor;
							}
							catch(Throwable e) {
							}
						}

						// code to get the style for the cell goes here
						cellStyle.setFillForegroundColor(color.getIndex());

					} else if  (cellStyle instanceof XSSFCellStyle)
						((XSSFCellStyle)cellStyle).setFillForegroundColor(new XSSFColor(new java.awt.Color(colorRGB[0], colorRGB[1], colorRGB[2])));
				}
				cellStyles.put(style,cellStyle);
			}
			cell.setCellStyle(cellStyle);
		}
		
	}

	private static Integer nullToZero(Integer integer) {
		return integer == null ? new Integer(0) : integer;
	}

	
	/**
	 * @param wb
	 * @param headerRowNum
	 * @param readRowNum
	 * @param metaDataHeaders
	 * @param outputHeaders
	 * @param parser
	 * @return
	 * @throws IOException
	 * @throws  
	 */
	private static void parseRowMessage(Workbook wb, int headerRowNum, int readRowNum,
			HashMap<String, Integer> metaDataHeaders, HashMap<String, Integer> outputHeaders, MessageParser parser, boolean isLast)
			throws IOException {
		String messageStr = getCellString(wb, 0, readRowNum, 0);
		if (!messageStr.isEmpty()) {

			PostData inputData = generateInitialePostData(wb, readRowNum, metaDataHeaders);
			inputData.setOriginalMessageText(messageStr);
			if (metaDataHeaders.get("Comments") != null)
			{
				String comments = getCellString(wb, 0, readRowNum, metaDataHeaders.get("Comments"));
				if (comments.isEmpty() == false)
				{
					inputData.setComments(new ArrayList<Comment>(Arrays.asList(new Comment(null,inputData.getFieldValueString(PostFieldType.PostPublisherId),comments, null, true))));
				}
			}
			
			
			/*inputData = new PostData();
			inputData.addField(PostFieldType.PostPublisherId, "582793141782660");
			inputData.addField(PostFieldType.ForumCategory, "RealEstate");
			inputData.addField(PostFieldType.ForumLanguage, "he");
			inputData.addField(PostFieldType.ForumCurrency, "NIS");
			inputData.addField(PostFieldType.ForumContractType, "Rent");
			inputData.addField(PostFieldType.ForumLocationCountry, "IL");
			inputData.addField(PostFieldType.ForumLocationCityCode, "3000");
			inputData.addField(PostFieldType.BillboardId, VariantEnum.generateInstance(VariantTypeEnums.Billboard.Facebook));
			inputData.addField(PostFieldType.PostTimeCreated, VariantDate.generateInstance(1, 1, 2000));
			
			inputData.setOriginalMessageText("לכל תושבי נחלאות / דירות קטנות\n" +
"שולחן מאיקאה במצב מעולה!!!    \n" +
"יעיל , מתקפל ומאוד נוח לשימוש ואחסון \n" +
"מחיר 500 ש״ח \n" +
"מידות : \n" +
"מצב סגור- 26 על 89 ס״מ \n" +
"חצי פתוח- 89 על 85  ס״מ\n" +
"פתוח - 89 על 144 ס״מ \n" +
"איסוף מנחלאות\n");*/
			if (workThreaded)
			{
				try {
					threadInputQueue.put(new ParseThreadData(readRowNum,inputData, null));
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
				executorService.execute(new Runnable() {
					@Override
					public void run() {
						DataRunner.parseMessageThreaded();
						
					}
				});
			}
			else 
			{
				PostData outData = parser.parseMessage(inputData, true);
				threadInputQueue.add(new ParseThreadData(readRowNum,inputData, outData));
			}
		}

		if (isLast)
		{
			executorService.shutdown();
			try {
				executorService.awaitTermination(1, TimeUnit.MINUTES);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
		
		while (threadOutputQueue.isEmpty() == false)
		{
			ParseThreadData data = threadOutputQueue.poll();
			
			gatherColumnsFromOutputData(outputHeaders, data.getInput(), data.getOutput(), wb, headerRowNum, metaDataHeaders.size() + 2);

			// Write the data
			Iterator<Entry<String, Integer>> headerIt = outputHeaders.entrySet().iterator();
			while (headerIt.hasNext()) {
				Entry<String, Integer> next = headerIt.next();
				String fieldName = next.getKey();
				String value = getFieldSafe(data.getOutput(), fieldName);
				getCell(wb, 0, data.getRow(), next.getValue(), true).setCellValue(value);
			}
		}
	}

	protected static void parseMessageThreaded() {
		MessageParser parser = threadMessageParser.get();
		if (parser == null)
		{
			ApiProxy.setEnvironmentForCurrentThread(testEnvironment);

			parser = new MessageParser();
			if (connectToLocalServer)
				try {
					connectToLocalServer = connectToLocalServer();
				} catch (IOException e) {
					e.printStackTrace();
				}
			parser.init(connectToLocalServer);
			threadMessageParser.set(parser);
		}
		ParseThreadData poll = threadInputQueue.poll();
		PostData outData = parser.parseMessage(poll.getInput(), true);
		try {
			threadOutputQueue.put(new ParseThreadData(poll.getRow(), poll.getInput(), outData));
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	/**
	 * @param target
	 * @param wb
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private static void saveWorkbook(Workbook wb, File target) throws FileNotFoundException, IOException {
		boolean setReadOnly = (target.exists() && target.canWrite()) ? false : true;
		// try to save to a temporary file than rename
		String tempFile = extendFileName(target.getAbsolutePath(), ".tmp");
		FileOutputStream out = new FileOutputStream(tempFile);
		wb.write(out);
		out.close();

		target.delete();
		if (new File(tempFile).renameTo(target)) {
			if (setReadOnly)
				target.setReadOnly();
		}
	}

	/*
	 * private static void copyCellProps(Workbook workbook, int rowNum, int colNumSource, int colNumTarget) { Cell
	 * srcCell = getCell(workbook, rowNum, colNumSource, false); if (srcCell != null) { Cell trgCell = getCell(workbook,
	 * rowNum, colNumTarget, true); trgCell.setCellStyle(srcCell.getCellStyle()); } }
	 */

	private static String getCellString(Workbook source, int sheetNum, int rowNum, int columnNum) {
		Cell cell = getCell(source, sheetNum, rowNum, columnNum, false);
		return getCellValueAsString(cell);
	}

	private static String getCellValueAsString(Cell cell) {
		if (cell != null) {
			int resType = cell.getCellType();
			if (resType == Cell.CELL_TYPE_FORMULA)
				resType = cell.getCachedFormulaResultType();
			switch (resType) {
			case Cell.CELL_TYPE_NUMERIC:
				double val = cell.getNumericCellValue();
				if (val - Math.floor(val) == 0)
					return Long.toString((long) val);
				else
					return Double.toString(val);

			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
				// case Cell.CELL_TYPE_FORMULA: return cell.getCellFormula();
			case Cell.CELL_TYPE_BOOLEAN:
				return Boolean.toString(cell.getBooleanCellValue());
			default:
				break;
			}
		}
		return "";
	}

	private static boolean isEndRow(Workbook wb, int sheetNum, int rowNum) {
		Sheet sheet = wb.getSheetAt(sheetNum);
		return (sheet == null) || (rowNum > sheet.getLastRowNum());
	}

	private static int getRowWidth(Workbook wb, int sheetNum, int rowNum) {
		Sheet sheet = wb.getSheetAt(sheetNum);
		if (sheet != null) {
			Row row = sheet.getRow(rowNum);
			if (row != null)
				return row.getLastCellNum();
		}
		return 0;
	}

	private static Cell getCell(Workbook source, int sheetNum, int rowNum, int columnNum, boolean createIfNeeded) {
		Sheet sheet = source.getSheetAt(sheetNum);
		if ((sheet == null) && (createIfNeeded) && (sheetNum == 0))
			sheet = source.createSheet();
		else if ((sheet == null) && (sheetNum != 0))
			throw new IllegalArgumentException("Unknown sheet number");
		if (sheet != null) {
			Row row = sheet.getRow(rowNum);
			if ((row == null) && (createIfNeeded))
				row = sheet.createRow(rowNum);
			if (row != null) {
				Cell cell = row.getCell(columnNum);
				if ((cell == null) && (createIfNeeded))
					cell = row.createCell(columnNum);
				return cell;
			}
		}
		return null;
	}

	private static Workbook openXLS(File source) throws Exception {
		Workbook exWorkBook = WorkbookFactory.create(new FileInputStream(source));
		// OPCPackage pkg = OPCPackage.open(new FileInputStream(source));
		// new Workbook(new FileInputStream(source))
		return exWorkBook;

	}

	/**
	 * @param expectedHeaders
	 * @param outData
	 * @param headerRowNum
	 * @param wb
	 * @param headerOffset
	 */
	private static void gatherColumnsFromOutputData(Map<String, Integer> expectedHeaders, PostData inData,
			PostData outData, Workbook wb, int headerRowNum, int expectedHeadersOffset) {
		if (outData != null) {
			Iterable<String> itr = outData.getFieldNamesSet();
			for(String ent : itr) {
				if ((expectedHeaders.containsKey(ent) == false) &&
						(inData.getField(ent) == null)) {
					Cell headerCell = getCell(wb, 0, headerRowNum, expectedHeadersOffset + expectedHeaders.size(), true);
					if (headerCell != null)
						headerCell.setCellValue(ent);
					expectedHeaders.put(ent, expectedHeadersOffset + expectedHeaders.size());
				}
			}
		}
	}

	private static PostData generateInitialePostData(Workbook wb, int readRowNum, Map<String, Integer> metaDataHeaders)
			throws IOException {
		PostData postData = basePostData.shalowClone();
		if (!postData.hasDefinedValue(PostFieldType.PostPublisherId))
			postData.addField(PostFieldType.PostPublisherId, "TestPublisher");
		if (!postData.hasDefinedValue(PostFieldType.BillboardId))
			postData.addField(PostFieldType.BillboardId,
					VariantEnum.generateInstance(VariantTypeEnums.Billboard.Facebook));
		if (!postData.hasDefinedValue(PostFieldType.ForumLanguage))
			postData.addField(PostFieldType.ForumLanguage, VariantEnum.generateInstance(VariantTypeEnums.Language.en));
		if (!postData.hasDefinedValue(PostFieldType.ForumCurrency))
			postData.addField(PostFieldType.ForumCurrency, VariantEnum.generateInstance(VariantTypeEnums.Currency.NIS));
		if (!postData.hasDefinedValue(PostFieldType.ForumLocationCountry))
			postData.addField(PostFieldType.ForumLocationCountry,
					VariantEnum.generateInstance(VariantTypeEnums.Country.IL));
		if (!postData.hasDefinedValue(PostFieldType.ForumPricePeriod))
			postData.addField(PostFieldType.ForumPricePeriod,
					VariantEnum.generateInstance(VariantTypeEnums.Period.Monthly));
		if (!postData.hasDefinedValue(PostFieldType.PostTimeCreated))
			postData.addField(PostFieldType.PostTimeCreated, VariantDate.generateInstance(1, 1, 2000));

		Iterator<Entry<String, Integer>> metaDataIt = metaDataHeaders.entrySet().iterator();
		while (metaDataIt.hasNext()) {
			Entry<String, Integer> next = metaDataIt.next();
			String fieldName = next.getKey();
			if (isOrigMessageProperty(fieldName)) {
				try {
					String val = getCellString(wb, 0, readRowNum, next.getValue());
					if (val.isEmpty() == false)
						postData.addField(fieldName, null, val);
				} catch (Exception e) {
				}
			}

		}

		return postData;
	}

	/**
	 * @param fieldName
	 */
	private static boolean isOrigMessageProperty(String fieldName) {
		return (fieldName.startsWith("Forum") || fieldName.equals("BillboardId") || fieldName.equals("PostId") ||
				fieldName.equals("PostTimeCreated") || fieldName.equals("PostTimeUpdated") || fieldName
					.equals("PostPublisherId"));
	}

	static private String getFieldSafe(PostData outData, String name) {
		if (outData != null) {
			Variant data = outData.getField(name);
			if (data != null)
			{
				if (data instanceof VariantDate)
					return data.toString();
				else return data.getStringValue();
			}
		}
		return "";
	}

}
