package Utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil 
{
	Workbook wb;
	//constructor for reading excel path
	public ExcelFileUtil(String Excelpath) throws Throwable
	{
		FileInputStream fi = new FileInputStream(Excelpath);
		wb = WorkbookFactory.create(fi);
	}
	
	//method for counting no of rows in a sheet
	public int rowCount(String sheetName)
	{
		return wb.getSheet(sheetName).getLastRowNum();
	}
	
	//method for reading cell data
	public String getCellData(String sheetName,int row,int col)
	{
		String data = "";
		if(wb.getSheet(sheetName).getRow(row).getCell(col).getCellType()==CellType.NUMERIC)
		{
			int celldata = (int) wb.getSheet(sheetName).getRow(row).getCell(col).getNumericCellValue();
			data = String.valueOf(celldata);
		}else
		{
			data = wb.getSheet(sheetName).getRow(row).getCell(col).getStringCellValue();
		}
		return data;
	}
	
	//method for writing status into new wb
	public void setCellData(String sheetName,int row,int col,String status,String WriteExcelPath) throws Throwable
	{
		//get sheet from wb
		Sheet ws = wb.getSheet(sheetName);
		//get row from sheet
		Row rownum = ws.getRow(row);
		//create cell
		Cell cell = rownum.createCell(col);
		//write status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(col).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Fail"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(col).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(col).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(WriteExcelPath);
		wb.write(fo);
		
	}
	
	public static void main(String[] args) throws Throwable 
	{
		ExcelFileUtil xl = new ExcelFileUtil("D:\\Live_Project_Automation\\Sample_book.xlsx");
		//count rows in emp sheet
		int rc = xl.rowCount("Emp");
		System.out.println(rc);
		for(int i=1;i<=rc;i++)
		{
			String FirstName = xl.getCellData("Emp", i, 0);
			String MiddleName = xl.getCellData("Emp", i, 1);
			String LastName = xl.getCellData("Emp", i, 2);
			String Eid = xl.getCellData("Emp", i, 3);
			System.out.println(FirstName+"     "+MiddleName+"      "+LastName+"      "+Eid);
			//xl.setCellData("Emp", i, 4, "pass", "D:\\Live_Project_Automation\\Results.xlsx");
			//xl.setCellData("Emp", i, 4, "fail", "D:\\Live_Project_Automation\\Results.xlsx");
			xl.setCellData("Emp", i, 4, "blocked", "D:\\Live_Project_Automation\\Results.xlsx");
		}
	}
}
