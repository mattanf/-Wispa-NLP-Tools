import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;

import java.util.logging.Logger;

import com.google.appengine.tools.development.testing.LocalDatastoreServiceTestConfig;
import com.google.appengine.tools.development.testing.LocalServiceTestHelper;
import com.pairapp.engine.parser.MessageParser;
import com.pairapp.engine.parser.data.PostData;
import com.pairapp.engine.parser.data.PostFieldData;
import com.pairapp.engine.parser.data.PostFieldType;
import com.pairapp.engine.parser.data.VariantDate;
import com.pairapp.engine.parser.data.VariantEnum;
import com.pairapp.engine.parser.data.VariantTypeEnums;
import com.pairapp.utilities.LogLineFormatter;

public class DataRunner {

	private final static String SUBJECT_MESSAGE = "Message";
	private final static String SUBJECT_RESULT = "Result";
	private final static String SUBJECT_FAILED_FIELDS = "Failed Fields";
	private final static int COMPARISON_FIELD_COUNT = 2;
	private final static int MAX_HEADERS = 100;
	private final static String SUBJECT_OPEN_ENDED_OUTPUTS = "...";

	private final static String KEY_LOG_LOCATIONS = "Log locations";
	private final static String KEY_PROCESS_LOCATIONS = "Process locations";
	private static PostData basePostData;
	private static final LocalServiceTestHelper datastoreHelper = new LocalServiceTestHelper(
			new LocalDatastoreServiceTestConfig());

	private static int startParseRow = -1;
	private static int endParseRow = 10000000;

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

	@SuppressWarnings("unchecked")
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
			while (!getCellString(wb, readRowNum, 0).equalsIgnoreCase(SUBJECT_MESSAGE) && isSuccess) {
				String h1 = getCellString(wb, readRowNum, 0).trim();
				String h2 = getCellString(wb, readRowNum, 1);

				if ((!h1.isEmpty()) && (!h2.isEmpty()) && (!h1.contains(" "))) {
					basePostData.addField(h1, null, h2);
				}
				++readRowNum;
				isSuccess = !isEndRow(wb, readRowNum);
			}
			if (!isSuccess)
				System.err.println("Error: ParserTree could not be find header message.");
			else {
				isSuccess = parser.init(false);
				if (!isSuccess)
					System.err.println("Error: ParserTree could not be initialized. Data source could not be found.");
				else {
					int metaDataHeadersOffset = 0;
					ArrayList<String> metaDataHeaders = new ArrayList<String>();
					int comparisonHeadersOffset = 0;
					ArrayList<String> comparisonHeaders = new ArrayList<String>();
					int outputHeadersOffset = 0;
					ArrayList<String> outputHeaders = new ArrayList<String>();
					boolean hasComparison = false;
					boolean isOutputHeaderFixed = false;

					// Read the header
					int sectionInt = 0;
					int rowWidth = getRowWidth(wb, readRowNum);
					int headerRowNum = readRowNum;
					for (int i = 0; i < rowWidth; ++i) {
						String columnName = getCellString(wb, readRowNum, i).trim();
						if (columnName.matches("Column\\d+"))
							columnName = "";
						switch (sectionInt) {
						case 0:
							if ((columnName.isEmpty() == false) &&
									(columnName.compareToIgnoreCase(SUBJECT_MESSAGE) != 0))
								metaDataHeaders.add(columnName);
							if (columnName.isEmpty() == true)
								sectionInt = 1;
							break;
						case 1:
							if (columnName.isEmpty() == false) {
								hasComparison = true;
								comparisonHeaders.add(columnName);
							} else
								sectionInt = 2;
							break;
						case 2:
							if (columnName.compareToIgnoreCase(SUBJECT_OPEN_ENDED_OUTPUTS) == 0)
								isOutputHeaderFixed = false;
							else if ((columnName.isEmpty() == false) &&
									(columnName.compareToIgnoreCase(SUBJECT_RESULT) != 0) &&
									(columnName.compareToIgnoreCase(SUBJECT_FAILED_FIELDS) != 0)) {
								outputHeaders.add(columnName);
								isOutputHeaderFixed = true;
							} else
								sectionInt = 3;
							break;
						}
					}
					++readRowNum;

					metaDataHeadersOffset = 1;
					comparisonHeadersOffset = metaDataHeaders.size() + 1 + metaDataHeadersOffset;
					outputHeadersOffset = comparisonHeaders.size() + 1 + comparisonHeadersOffset;
					int countParsedMessages = 0;

					if (outputHeaders.size() == 0)
						outputHeaders = (ArrayList<String>) (comparisonHeaders.clone());

					if (hasComparison) {
						outputHeaders.add(0, SUBJECT_RESULT);
						outputHeaders.add(1, SUBJECT_FAILED_FIELDS);
					}

					int countMessageDiffer[] = new int[MAX_HEADERS];
					Arrays.fill(countMessageDiffer, 0);

					readRowNum = Math.max(readRowNum, startParseRow);

					// Compare and write the report
					while (!isEndRow(wb, readRowNum)) {
						if (readRowNum > endParseRow)
							break;
						String messageStr = getCellString(wb, readRowNum, 0);
						if (!messageStr.isEmpty()) {

							PostData inputData = generateInitialePostData(wb, readRowNum, metaDataHeaders,
									metaDataHeadersOffset);

							inputData.setOriginalMessageText(messageStr);
							PostData outData = parser.parseMessage(inputData);
							
							++countParsedMessages;
							if (!isOutputHeaderFixed)
								gatherColumnsFromOutputData(outputHeaders, inputData, outData, wb, headerRowNum,
										outputHeadersOffset);

							String resultDesc = "";
							String failedFieldsDesc = "";
							if ((outData != null) && (hasComparison == true)) {
								resultDesc = "Same";

								for (int i = COMPARISON_FIELD_COUNT; i < comparisonHeaders.size(); ++i) {
									String outputKeyName = comparisonHeaders.get(i);

									int col = getColumnNum(comparisonHeadersOffset, comparisonHeaders, outputKeyName);
									if (getCellString(wb, readRowNum, col).compareToIgnoreCase(
											getFieldSafe(outData, outputKeyName)) != 0) {
										resultDesc = "Different";
										if (failedFieldsDesc.isEmpty() == false)
											failedFieldsDesc += "+";
										else
											++countMessageDiffer[0];
										failedFieldsDesc = failedFieldsDesc + outputKeyName;
										++countMessageDiffer[i];
									}
								}
							}

							// Write the data
							// Write the comparison headers
							writeToCell(wb, readRowNum, outputHeadersOffset, outputHeaders, SUBJECT_RESULT, resultDesc);
							writeToCell(wb, readRowNum, outputHeadersOffset, outputHeaders, SUBJECT_FAILED_FIELDS,
									failedFieldsDesc);

							// Output
							for (int i = hasComparison ? COMPARISON_FIELD_COUNT : 0; i < outputHeaders.size(); ++i) {
								String fieldName = outputHeaders.get(i);
								String value = getFieldSafe(outData, fieldName);
								writeToCell(wb, readRowNum, outputHeadersOffset, outputHeaders, fieldName, value);
							}
						}
						++readRowNum;
	
					}
					
					// Write the comparison results
						//if (hasComparison)
						//{
						//	writeComparisonResults(wb, countParsedMessages,comparisonHeaders, countMessageDiffer);
						//}
					saveWorkbook(wb, target);
				}
			}
		} catch (IOException e) {
			System.err.println("Error: " + e.getMessage());

			isSuccess = false;
		}

		return isSuccess;
	}

	private static void writeComparisonResults(Workbook wb, int countParsedMessages,
			ArrayList<String> comparisonHeaders, int[] countMessageDiffer) {
		// TODO Auto-generated method stub

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

	/**
	 * @param wb
	 * @param rowNum
	 * @param outputHeadersOffset
	 * @param outputHeaders
	 * @param failedFieldsDesc
	 */
	private static void writeToCell(Workbook wb, int rowNum, int outputHeadersOffset,
			ArrayList<String> outputHeaders, String fieldName, String failedFieldsDesc) {
		int colNum;
		colNum = getColumnNum(outputHeadersOffset, outputHeaders, fieldName);
		if (colNum != -1) {
			Cell cell = getCell(wb, rowNum, colNum, true);
			if (cell != null)
				cell.setCellValue(failedFieldsDesc);
		}
	}

	private static int getColumnNum(int offset, ArrayList<String> array, String key) {
		int index = array.indexOf(key);
		return index == -1 ? -1 : index + offset;
	}

	/*
	 * private static void copyCellProps(Workbook workbook, int rowNum, int colNumSource, int colNumTarget) {
	 * Cell srcCell = getCell(workbook, rowNum, colNumSource, false); if (srcCell != null) { Cell trgCell =
	 * getCell(workbook, rowNum, colNumTarget, true); trgCell.setCellStyle(srcCell.getCellStyle()); } }
	 */

	private static String getCellString(Workbook source, int rowNum, int columnNum) {
		Cell cell = getCell(source, rowNum, columnNum, false);
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
					return Long.toString((int) val);
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

	private static boolean isEndRow(Workbook wb, int rowNum) {
		Sheet sheet = wb.getSheetAt(0);
		return (sheet == null) || (rowNum > sheet.getLastRowNum());
	}

	private static int getRowWidth(Workbook wb, int rowNum) {
		Sheet sheet = wb.getSheetAt(0);
		if (sheet != null) {
			Row row = sheet.getRow(rowNum);
			if (row != null)
				return row.getLastCellNum();
		}
		return 0;
	}

	private static Cell getCell(Workbook source, int rowNum, int columnNum, boolean createIfNeeded) {
		Sheet sheet = source.getSheetAt(0);
		if ((sheet == null) && (createIfNeeded))
			sheet = source.createSheet();
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
		//OPCPackage pkg = OPCPackage.open(new FileInputStream(source));
		//new Workbook(new FileInputStream(source))
		return exWorkBook;

	}

	/**
	 * @param expectedHeaders
	 * @param outData
	 * @param headerRowNum
	 * @param wb
	 * @param headerOffset
	 */
	private static void gatherColumnsFromOutputData(ArrayList<String> expectedHeaders, PostData inData,
			PostData outData, Workbook wb, int headerRowNum, int headerOffset) {
		if (outData != null) {
			Iterator<PostFieldData> it = outData.getFieldIterator();
			while (it.hasNext()) {
				PostFieldData field = it.next();
				if ((expectedHeaders.contains(field.getName()) == false) && (inData.getField(field.getName()) == null)) {
					Cell headerCell = getCell(wb, headerRowNum, headerOffset + expectedHeaders.size(), true);
					if (headerCell != null)
						headerCell.setCellValue(field.getName());
					expectedHeaders.add(field.getName());
				}
			}
		}
	}

	private static PostData generateInitialePostData(Workbook wb, int readRowNum,
			ArrayList<String> metaDataHeaders, int metaDataHeadersOffset) throws IOException {
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

		for (int i = 0; i < metaDataHeaders.size(); ++i) {
			String fieldName = metaDataHeaders.get(i);
			if (fieldName.startsWith("Forum") || fieldName.equals("BillboardId") || fieldName.equals("PostId") ||
					fieldName.equals("PostTimeCreated") || fieldName.equals("PostTimeUpdated") ||
					fieldName.equals("PostPublisherId")) {
				try {
					String val = getCellString(wb, readRowNum, i + metaDataHeadersOffset);
					postData.addField(metaDataHeaders.get(i), null, val);
				} catch (Exception e) {
				}
			}

		}

		return postData;
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
