package utility;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

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

	public static boolean saveXLS(XSSFWorkbook wb, String fileName) {
		try {
			FileOutputStream fileOut = new FileOutputStream(fileName + ".tmp");
			//write this workbook to an Outputstream.
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			wb.getPackage().close();
			new File(fileName).delete();
			new File(fileName + ".tmp").renameTo(new File(fileName));
			
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	public static int getSheetNumber(XSSFWorkbook source, String name) 
	{
		return source != null ? source.getSheetIndex(name) : -1;
	}
	
	public static String getCellString(XSSFWorkbook source, int rowNum, int columnNum) 
	{
		return getCellString(source, 0, rowNum, columnNum);
	}

	public static String getCellString(XSSFWorkbook source, int sheetNum, int rowNum, int columnNum)
	{
		XSSFCell cell = getCell(source, sheetNum, rowNum, columnNum);
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
		return isEndRow(wb, 0, rowNum);
	}
	public static boolean isEndRow(XSSFWorkbook wb, int sheetNum, int rowNum) {
		XSSFSheet sheet = wb.getSheetAt(sheetNum);
		return (sheet == null) || (rowNum > sheet.getLastRowNum());
	}

	public static XSSFCell getCell(XSSFWorkbook source, int rowNum, int columnNum)
	{
		return getCell(source, 0, rowNum, columnNum);
	}
	
	public static XSSFCell getCell(XSSFWorkbook source, int sheetNumber, int rowNum, int columnNum)
	{
		return getCell(source, sheetNumber, rowNum, columnNum, false);
	}
	public static XSSFCell getCell(XSSFWorkbook source, int sheetNumber, int rowNum, int columnNum, boolean createIfNeeded) {
		if (sheetNumber >= 0)
		{
			XSSFSheet sheet = source.getSheetAt(sheetNumber);
			if (sheet != null) {
				XSSFRow row = sheet.getRow(rowNum);
				if ((row == null) && (createIfNeeded))
					row = sheet.createRow(rowNum);
				if (row != null) {
					XSSFCell cell =  row.getCell(columnNum);
					if ((cell == null) && (createIfNeeded))
						cell = row.createCell(columnNum);
					return cell;
				}
			}
		}
		
		return null;
	}

	public static boolean setCellString(XSSFWorkbook source, int rowNum, int columnNum, String string) {
		return setCellString(source, 0, rowNum, columnNum, string);
	}
	
	public static boolean setCellString(XSSFWorkbook source, int sheetNum, int rowNum, int columnNum, String string) {
		XSSFCell cell = getCell(source, sheetNum, rowNum, columnNum, true);
		if (cell != null)
		{
			cell.setCellValue(string);
		}
		return cell != null;
	}

}
