package utility;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSUtil {


	public static XSSFWorkbook openXLS(String source) {
		try {
			OPCPackage pkg = OPCPackage.open(new String(source));
			return new XSSFWorkbook(pkg);
		} catch (Exception e) {
			return null;
		}

	}

	public static String getCellString(XSSFWorkbook source, int rowNum, int columnNum) {
		XSSFCell cell = getCell(source, rowNum, columnNum);
		return getCellValueAsString(cell);
	}

	public static String getCellValueAsString(XSSFCell cell) {
		if (cell != null) {
			int resType = cell.getCellType();
			if (resType == XSSFCell.CELL_TYPE_FORMULA)
				resType = cell.getCachedFormulaResultType();
			switch (resType) {
			case XSSFCell.CELL_TYPE_NUMERIC:
				double val = cell.getNumericCellValue();
				if (val - Math.floor(val) == 0)
					return Long.toString((int) val);
				else
					return Double.toString(val);

			case XSSFCell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
				// case Cell.CELL_TYPE_FORMULA: return cell.getCellFormula();
			case XSSFCell.CELL_TYPE_BOOLEAN:
				return Boolean.toString(cell.getBooleanCellValue());
			default:
				break;
			}
		}
		return "";
	}

	public static int getRowWidth(XSSFWorkbook wb, int rowNum) {
		XSSFSheet sheet = wb.getSheetAt(0);
		if (sheet != null) {
			XSSFRow row = sheet.getRow(rowNum);
			if (row != null)
				return row.getLastCellNum();
		}
		return 0;
	}

	public static boolean isEndRow(XSSFWorkbook wb, int rowNum) {
		XSSFSheet sheet = wb.getSheetAt(0);
		return (sheet == null) || (rowNum > sheet.getLastRowNum());
	}

	public static XSSFCell getCell(XSSFWorkbook source, int rowNum, int columnNum) {
		XSSFSheet sheet = source.getSheetAt(0);
		if (sheet != null) {
			XSSFRow row = sheet.getRow(rowNum);
			if (row != null) {
				return row.getCell(columnNum);
			}
		}
		return null;
	}
}
