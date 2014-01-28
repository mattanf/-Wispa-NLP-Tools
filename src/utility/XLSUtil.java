package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class XLSUtil {

	public static Workbook openXLS(String source) {
		return openXLS(new File(source));
	}

	public static Workbook openXLS(File source) {
		try {
			Workbook exWorkBook = WorkbookFactory.create(new FileInputStream(source));
			return exWorkBook;
		} catch (Exception e) {
			return null;
		}
	}

	public static boolean saveXLS(Workbook wb, String fileName) {
		return saveXLS(wb, new File(fileName));
	}

	public static boolean saveXLS(Workbook wb, File fileName) {
		try {
			FileOutputStream fileOut = new FileOutputStream(fileName.getPath() + ".tmp");
			// write this workbook to an Outputstream.
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			fileName.delete();
			new File(fileName.getPath() + ".tmp").renameTo(fileName);

		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	public static int getSheetNumber(Workbook source, String name) {
		return source != null ? source.getSheetIndex(name) : -1;
	}

	public static String getCellString(Workbook source, int rowNum, int columnNum) {
		return getCellString(source, 0, rowNum, columnNum);
	}

	public static String getCellString(Workbook source, int sheetNum, int rowNum, int columnNum) {
		Cell cell = getCell(source, sheetNum, rowNum, columnNum);
		return getCellValueAsString(cell);
	}

	public static String getCellValueAsString(Cell cell) {
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
				String str = cell.getStringCellValue().trim();
				str = str.replaceAll("[\\u00a0]", " ");
				return str;
				// case Cell.CELL_TYPE_FORMULA: return cell.getCellFormula();
			case Cell.CELL_TYPE_BOOLEAN:
				return Boolean.toString(cell.getBooleanCellValue());
			default:
				break;
			}
		}
		return "";
	}

	public static int getRowWidth(Workbook wb, int rowNum) {
		Sheet sheet = wb.getSheetAt(0);
		if (sheet != null) {
			Row row = sheet.getRow(rowNum);
			if (row != null)
				return row.getLastCellNum();
		}
		return 0;
	}

	public static boolean isEndRow(Workbook wb, int rowNum) {
		return isEndRow(wb, 0, rowNum);
	}

	public static boolean isEndRow(Workbook wb, int sheetNum, int rowNum) {
		Sheet sheet = wb.getSheetAt(sheetNum);
		return (sheet == null) || (rowNum > sheet.getLastRowNum());
	}

	public static Cell getCell(Workbook source, int rowNum, int columnNum) {
		return getCell(source, 0, rowNum, columnNum);
	}

	public static Cell getCell(Workbook source, int sheetNumber, int rowNum, int columnNum) {
		return getCell(source, sheetNumber, rowNum, columnNum, false);
	}

	public static Cell getCell(Workbook source, int sheetNumber, int rowNum, int columnNum, boolean createIfNeeded) {
		if (sheetNumber >= 0) {
			Sheet sheet = source.getSheetAt(sheetNumber);
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
		}

		return null;
	}

	public static boolean setCellString(Workbook source, int rowNum, int columnNum, String string) {
		return setCellString(source, 0, rowNum, columnNum, string);
	}

	public static boolean setCellString(Workbook source, int sheetNum, int rowNum, int columnNum, String string) {
		Cell cell = getCell(source, sheetNum, rowNum, columnNum, true);
		if (cell != null) {
			cell.setCellValue(string);
		}
		return cell != null;
	}

	public static boolean setCellValue(Workbook source, int sheetNum, int rowNum, int columnNum, Boolean value) {
		Cell cell = getCell(source, sheetNum, rowNum, columnNum, true);
		if (cell != null) {
			if (value == null)
				cell.setCellValue("");
			else
				cell.setCellValue(value);
		}
		return cell != null;
	}

	public static boolean setCellValue(Workbook source, int sheetNum, int rowNum, int columnNum, Integer value) {
		Cell cell = getCell(source, sheetNum, rowNum, columnNum, true);
		if (cell != null) {
			if (value == null)
				cell.setCellValue("");
			else
				cell.setCellValue(value);
		}
		return cell != null;
	}

	public static boolean setCellValue(Workbook source, int sheetNum, int rowNum, int columnNum, Double value) {
		Cell cell = getCell(source, sheetNum, rowNum, columnNum, true);
		if (cell != null) {
			if (value == null)
				cell.setCellValue("");
			else
				cell.setCellValue(value);
		}
		return cell != null;
	}

	// Get all the values in a row and enters them to a map with their column number
	public static Map<String, Integer> getHeader(Workbook wb, int sheetNum, int rowNum) {
		HashMap<String, Integer> map = new HashMap<>();

		Sheet sheet = wb.getSheetAt(sheetNum);
		if (sheet != null) {
			Row row = sheet.getRow(rowNum);
			if ((row != null) && (row.getLastCellNum() != -1)){
				// scan all available columns in the row
				for (int columnNum = row.getFirstCellNum(); columnNum <= row.getLastCellNum(); ++columnNum) {
					Cell cell = getCell(wb, sheetNum, rowNum, columnNum);
					String headerValue = getCellValueAsString(cell);
					if (headerValue.isEmpty() == false)
						map.put(headerValue, columnNum);
				}
			}
		}
		return map;
	}

	public static void updateHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> headerValues) {
		for (Entry<String, Integer> entry : headerValues.entrySet()) {
			int columnNum = entry.getValue();
			String name = entry.getKey();
			if (!XLSUtil.getCellString(wb, sheetNum, rowNum, columnNum).equals(name))
				XLSUtil.setCellString(wb, sheetNum, rowNum, columnNum, name);
		}

	}

	public static int getOrCreateSheet(Workbook wb, String sheetName, int defaultOrder) {
		Sheet sh = wb.getSheet(sheetName);
		if (sh == null) {
			wb.createSheet(sheetName);
			if (defaultOrder != -1)
				wb.setSheetOrder(sheetName, defaultOrder);
		}
		return wb.getSheetIndex(sheetName);
	}

	public static void orderRows(Workbook wb, int sheetNum, int orderFrom, List<Integer> rowsNumbers) {
		Sheet sheet = wb.getSheetAt(sheetNum);
		if (sheet != null) {
			ArrayList<Row> rows = new ArrayList<>();
			for (Integer rowNum : rowsNumbers) {
				rows.add(sheet.getRow(rowNum));
			}

			for (Row row : rows) {
				if (row != null)
					row.setRowNum(orderFrom++);
			}
		}
	}
	
	public static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, String value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellString(wb, sheetNum, rowNum, columnIndex, value);
	}
	
	public static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, Boolean value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}
	public static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, Integer value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}
	public static void setCellByHeader(Workbook wb, int sheetNum, int rowNum, Map<String, Integer> workHeader,
			String columnName, Double value) {
		Integer columnIndex = addColumnToHeader(workHeader, columnName);
		XLSUtil.setCellValue(wb, sheetNum, rowNum, columnIndex, value);
	}
	
	/**
	 * @param workHeader
	 * @param columnName
	 * @return
	 */
	public static Integer addColumnToHeader(Map<String, Integer> workHeader, String columnName) {
		Integer columnIndex = workHeader.get(columnName);
		if (columnIndex == null) {
			columnIndex = new Integer(0);
			for (Integer occupiedColumn : workHeader.values()) {
				columnIndex = Math.max(occupiedColumn + 1, columnIndex);
			}
			workHeader.put(columnName, columnIndex);
		}
		return columnIndex;
	}
	
	public static int getHeaderRow(Workbook wb, int sheetNum, Map<String, Integer> headerColumns, String ... neededColumns)
	{
		int rowNum = -1;
		boolean isValidHeader = false;
		do {
			++rowNum;
			Map<String, Integer> tmpColumns = XLSUtil.getHeader(wb, sheetNum, rowNum);
			isValidHeader = true;
			for(String neededColumn : neededColumns)
			{	
				if ((neededColumn != null) && (neededColumn.isEmpty() == false))
				{	
					isValidHeader &= tmpColumns.get(neededColumn) != null;
					if (isValidHeader)
						headerColumns.putAll(tmpColumns);
				}
			}
		} while ((isValidHeader == false) && (XLSUtil.isEndRow(wb, sheetNum, rowNum) == false));
		
		if (isValidHeader == true)
			return rowNum;
		else return -1;
	}
	

	public static String getPreHeaderProperty(Workbook wb, int sheetNum, int headerRow, String propertyName) {
		int curRow = 0;
		while ((curRow < headerRow) && (!XLSUtil.isEndRow(wb, sheetNum, curRow)))
		{
			String curPropName = XLSUtil.getCellString(wb, sheetNum, curRow, 0);
			if (propertyName.equalsIgnoreCase(curPropName))
			{
				return XLSUtil.getCellString(wb, sheetNum, curRow, 1);
			}
			++curRow;
		}
		return null;
	}
	
}
