package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil {
	XSSFWorkbook wb;
	//constructor to read excel file
	public ExcelFileUtil(String Excelpath) throws Throwable {
		FileInputStream fi = new FileInputStream(Excelpath);
		wb = new XSSFWorkbook(fi);
	}
	// count the no of rows
	public int rowCount(String Sheetname) {
		return wb.getSheet(Sheetname).getLastRowNum();
	}
	//method for reading cell data
	public String getCellData(String sheetname, int row, int column) {
		String data = "";
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC) {
			int celldata = (int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
			data = String.valueOf(celldata);
		}
		else {
			data = wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}
	public void setCellData(String sheetname, int row, int column, String status, String writeExcel) throws Throwable {
		XSSFSheet ws = wb.getSheet(sheetname);
		XSSFRow rowNum = ws.getRow(row);
		XSSFCell cell = rowNum.createCell(column);
		//write cell status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("pass")) {
			XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);	
			style.setFont(font);
			rowNum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Fail")) {
			XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);	
			style.setFont(font);
			rowNum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked")) {
			XSSFCellStyle style = wb.createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);	
			style.setFont(font);
			rowNum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(writeExcel);
				wb.write(fo);
	}
	public static void main(String[] args) throws Throwable {
		ExcelFileUtil xl = new ExcelFileUtil("D:\\RahmanShaik.xlsx");
		int rc = xl.rowCount("Empdata");
		System.out.println(rc);
		for(int i=0;i<=rc;i++) {
			String fname = xl.getCellData("Empdata", i, 0);
			String mname = xl.getCellData("Empdata", i, 1);
			String lname = xl.getCellData("Empdata", i, 2);
			String eid = xl.getCellData("Empdata", i, 3);
			System.out.println(fname+" "+mname+" "+lname+" "+eid);
			xl.setCellData("Empdata", i, 4, "Pass", "D://Reslelts.xlsx");
		}
	}
}
