import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;

import java.util.logging.Logger;

import com.google.appengine.tools.development.testing.LocalDatastoreServiceTestConfig;
import com.google.appengine.tools.development.testing.LocalServiceTestHelper;
import com.google.gwt.thirdparty.javascript.jscomp.graph.GraphColoring.Color;
import com.pairapp.engine.parser.MessageParser;
import com.pairapp.engine.parser.data.PostData;
import com.pairapp.engine.parser.data.PostFieldData;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.VariantDate;
import com.pairapp.engine.parser.data.VariantEnum;
import com.pairapp.engine.parser.data.VariantTypeEnums;
import com.pairapp.utilities.LogLineFormatter;

public class DataRunner {

	final static int STYLE_CENTER = 1 << 0;
	final static int STYLE_BORDER_RIGHT = 1 << 1;
	final static int STYLE_BORDER_LEFT = 1 << 2;
	final static int STYLE_BORDER_BOTTOM = 1 << 3;
	final static int STYLE_PERCENTAGE = 1 << 4;
	
	final static int STYLE_COLOR_GREEN = 1 << 16;
	final static int STYLE_COLOR_YELLOW = 1 << 17;
	final static int STYLE_COLOR_RED = 1 << 18;
	final static int STYLE_COLOR = STYLE_COLOR_GREEN | STYLE_COLOR_YELLOW | STYLE_COLOR_RED;
	final static int STYLE_COLOR_LIGHT = 1 << 19;
	
	private final static String SUBJECT_MESSAGE = "Message";
	private final static int MAX_HEADERS = 100;

	private static PostData basePostData;
	private static final LocalServiceTestHelper datastoreHelper = new LocalServiceTestHelper(
			new LocalDatastoreServiceTestConfig());

	private static int startParseRow = -1;
	private static int endParseRow = 10000000;

	private static HashMap<Integer, CellStyle> cellStyles;
	private static int usedColorIndexRunner = 0;

	// private static int endParseRow = -1;
	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {

		// This class does in a round about way initializes the log. we need to call it before making changes to the
		// loging mechnism
		// GeocodeQuerier.instance();
		// Remove the previous logging mechanism and add a better handler
		Logger.getGlobal().setUseParentHandlers(false);
		// getParent().removeHandler(Logger.getGlobal().getParent().getHandlers()[0]);
		ConsoleHandler handler = new ConsoleHandler();
		handler.setFormatter(new LogLineFormatter());
		Logger.getGlobal().addHandler(handler);
		Logger.getGlobal().setLevel(Level.SEVERE);

		int fileIndex = 0;
		for (int i = 0; i < args.length; ++i) {
			if (args[fileIndex].startsWith("-")) {
				if (args[fileIndex].matches("-r\\d+"))
					endParseRow = startParseRow = Integer.valueOf(args[fileIndex].substring(2)) - 1;
				if (args[fileIndex].matches("-rs\\d+"))
					startParseRow = Integer.valueOf(args[fileIndex].substring(3)) - 1;
				if (args[fileIndex].matches("-re\\d+"))
					endParseRow = Integer.valueOf(args[fileIndex].substring(3)) - 1;
				++fileIndex;
			} else
				break;
		}

		if (args.length - fileIndex == 0) {
			System.out.println("Usage: [XLS/X source file] {XLS/X output file}");
		} else {
			File sourceFile = new File(args[fileIndex]);

			if ((args[fileIndex].toLowerCase().endsWith(".xls") == false) &&
					(args[fileIndex].toLowerCase().endsWith(".xlsx") == false))
				System.err.println("Error: Specified input file '" + args[fileIndex] + "' is not a xls/x file.");
			if ((args.length > fileIndex + 1) &&
					((args[fileIndex + 1].toLowerCase().endsWith(".xls") == false) && (args[fileIndex + 1]
							.toLowerCase().endsWith(".xlsx") == false)))
				System.err.println("Error: Specified output file '" + args[fileIndex + 1] + "' is not a xls/x file.");
			else if (!sourceFile.isFile() || !sourceFile.exists())
				System.err.println("Error: Specified file '" + args[fileIndex] + "' could not be read.");
			else {
				String secondaryFileName = args[fileIndex];

				secondaryFileName = extendFileName(secondaryFileName, "-compiled");
				File targetFile = new File(secondaryFileName);
				if (args.length > fileIndex + 1)
					targetFile = new File(args[fileIndex + 1]);

				cellStyles = new HashMap<>();
				System.out.println("Started parse process");
				datastoreHelper.setUp();
				analyzeFile(sourceFile, targetFile);
				datastoreHelper.tearDown();
				System.out.println("Ended parse process");
			}
		}
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
				isSuccess = parser.init(false);
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
								Logger.getGlobal().info(
										"Parsing row " + (readRowNum - firstMessageRow) + " in stage " + (stage + 1));

							if (stage == 0)
								parseRowMessage(wb, headerRowNum, readRowNum, metaDataHeaders, outputHeaders, parser);
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
			if ((isOffer == isSeek) && (purposeString.isEmpty() == false))
				isSeek = false;
		Iterator<Entry<String, Integer>> outputHeaderIt = outputHeaders.entrySet().iterator();
			while (outputHeaderIt.hasNext()) {
				Entry<String, Integer> outputHeader = outputHeaderIt.next();
				String fieldName = outputHeader.getKey();
				Integer origHeaderPos = metaDataHeaders.get(fieldName);
				if (origHeaderPos != null) {
					
					// Compare the values
					String origValue = normNumeric(getCellString(wb, 0, readRowNum, origHeaderPos));
					String genValue = normNumeric(getCellString(wb, 0, readRowNum, outputHeader.getValue()));
					
					boolean isFalseNegative = (origValue.isEmpty() == false) && (genValue.isEmpty() == true);
					boolean isFalsePositive = !isFalseNegative && origValue.compareToIgnoreCase(genValue) != 0;
					boolean isSame = !isFalseNegative && !isFalsePositive;
					
					
					addComparisonData(fieldName, isValid, isOffer, isSeek, isFalseNegative, isFalsePositive, isSame,
							resultCount);
					setCellStyle(wb, 0, readRowNum, outputHeader.getValue(), (isFalseNegative ? STYLE_COLOR_YELLOW :
							(isFalsePositive ? STYLE_COLOR_RED : STYLE_COLOR_GREEN)) | (isValid ? 0 : STYLE_COLOR_LIGHT));
				}					
			}
		}
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

	private static void addComparisonData(String fieldName, boolean isValid, boolean isOffer, boolean isSeek,
			boolean isFalseNegative, boolean isFalsePositive, boolean isSame,
			TreeMap<String, HashMap<String, Integer>> resultCount) {
		HashMap<String, Integer> fieldMap = resultCount.get(fieldName);
		if (fieldMap == null) {
			fieldMap = new HashMap<>();
			resultCount.put(fieldName, fieldMap);
		}
		String baseKey = isFalseNegative ? "FN" : (isFalsePositive ? "FP" : "EQ");
		incrementFieldInMap(fieldMap, baseKey);
		if (isValid == true) {
			incrementFieldInMap(fieldMap, "ValidMessage-" + baseKey);
			if (isSeek == true)
				incrementFieldInMap(fieldMap, "Seek-ValidMessage-" + baseKey);
			if (isOffer == true)
				incrementFieldInMap(fieldMap, "Offer-ValidMessage-" + baseKey);
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

			Integer falseNegativeValidMessage = nullToZero(compData.getValue().get("ValidMessage-FN"));
			Integer falsePositiveValidMessage = nullToZero(compData.getValue().get("ValidMessage-FP"));
			Integer equalValidMessage = nullToZero(compData.getValue().get("ValidMessage-EQ"));
			Integer totalValidMessage = falseNegativeValidMessage + falsePositiveValidMessage + equalValidMessage;

			Integer falseNegativeValidMessageSeek = nullToZero(compData.getValue().get("Seek-ValidMessage-FN"));
			Integer falsePositiveValidMessageSeek = nullToZero(compData.getValue().get("Seek-ValidMessage-FP"));
			Integer equalValidMessageSeek = nullToZero(compData.getValue().get("Seek-ValidMessage-EQ"));
			Integer totalValidMessageSeek = falseNegativeValidMessageSeek + falsePositiveValidMessageSeek + equalValidMessageSeek;

			Integer falseNegativeValidMessageOffer = nullToZero(compData.getValue().get("Offer-ValidMessage-FN"));
			Integer falsePositiveValidMessageOffer = nullToZero(compData.getValue().get("Offer-ValidMessage-FP"));
			Integer equalValidMessageOffer = nullToZero(compData.getValue().get("Offer-ValidMessage-EQ"));
			Integer totalValidMessageOffer = falseNegativeValidMessageOffer + falsePositiveValidMessageOffer + equalValidMessageOffer;

			columnIndex = -1;
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellValue(fieldName);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(falseNegative + "/" + total);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(falsePositive + "/" + total);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true).setCellFormula(equal + "/" + total);

			getCell(wb, sheetNum, writeRow, ++columnIndex, true)
					.setCellFormula(falseNegativeValidMessage + "/" + totalValidMessage);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true)
					.setCellFormula(falsePositiveValidMessage + "/" + totalValidMessage);
			getCell(wb, sheetNum, writeRow, ++columnIndex, true)
					.setCellFormula(equalValidMessage + "/" + totalValidMessage);
			
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
					else //if ((style & STYLE_COLOR_RED) != 0)
						colorRGB = new short[] {255, 64, 64};
					
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
	 */
	private static void parseRowMessage(Workbook wb, int headerRowNum, int readRowNum,
			HashMap<String, Integer> metaDataHeaders, HashMap<String, Integer> outputHeaders, MessageParser parser)
			throws IOException {
		String messageStr = getCellString(wb, 0, readRowNum, 0);
		if (!messageStr.isEmpty()) {

			PostData inputData = generateInitialePostData(wb, readRowNum, metaDataHeaders);

			inputData.setOriginalMessageText(messageStr);
			PostData outData = parser.parseMessage(inputData);

			gatherColumnsFromOutputData(outputHeaders, inputData, outData, wb, headerRowNum, metaDataHeaders.size() + 2);

			// Write the data
			Iterator<Entry<String, Integer>> headerIt = outputHeaders.entrySet().iterator();
			while (headerIt.hasNext()) {
				Entry<String, Integer> next = headerIt.next();
				String fieldName = next.getKey();
				String value = getFieldSafe(outData, fieldName);
				getCell(wb, 0, readRowNum, next.getValue(), true).setCellValue(value);
			}
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
			Iterator<PostFieldData> it = outData.getFieldIterator();
			while (it.hasNext()) {
				PostFieldData field = it.next();
				if ((expectedHeaders.containsKey(field.getName()) == false) &&
						(inData.getField(field.getName()) == null)) {
					Cell headerCell = getCell(wb, 0, headerRowNum, expectedHeadersOffset + expectedHeaders.size(), true);
					if (headerCell != null)
						headerCell.setCellValue(field.getName());
					expectedHeaders.put(field.getName(), expectedHeadersOffset + expectedHeaders.size());
				}
			}
		}
	}

	private static PostData generateInitialePostData(Workbook wb, int readRowNum, Map<String, Integer> metaDataHeaders)
			throws IOException {
		PostData postData = basePostData.shalowClone();
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
			PostFieldData data = outData.getField(name);
			if (data != null)
				return data.getValue().getStringValue();
		}
		return "";
	}

}
