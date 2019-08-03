package com.samples.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {

	private static String[] bookHeaders= {"Book Id","Book Name","Author","Edition"};
	private static String[] employeeHeaders= {"Employee Id","Name","Username"};
	
	public static void main(String[] args) throws IOException, SQLException {
		
		Workbook workbook=new XSSFWorkbook();
		Sheet sheet_book=workbook.createSheet("Books");
		Sheet sheet_employee=workbook.createSheet("Employees");
		
		
		Font headerFont=workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 17);
		headerFont.setColor(IndexedColors.RED.getIndex());
		
		CellStyle headerCellStyle=workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		
		Row headerRowEmployee=sheet_employee.createRow(0);
		
		for (int i = 0; i < employeeHeaders.length; i++) {
			Cell cell=headerRowEmployee.createCell(i);
			cell.setCellValue(employeeHeaders[i]);
			cell.setCellStyle(headerCellStyle);
		}
		
		Row headerRowBook=sheet_book.createRow(0);
		
		for (int i = 0; i < bookHeaders.length; i++) {
			Cell cell=headerRowBook.createCell(i);
			cell.setCellValue(bookHeaders[i]);
			cell.setCellStyle(headerCellStyle);
		}
		DaoFile daoFile=new DaoFile();
		String book_query="Select * from Book";
		String employee_query="Select * from Employee";
		
		ResultSet rs=null;
		rs=daoFile.getResult(employee_query);
		int rowNumberEmployee=1;
		while(rs.next())
		{
			Row row=sheet_employee.createRow(rowNumberEmployee++);
			row.createCell(0).setCellValue(rs.getInt("EmpId"));
			row.createCell(1).setCellValue(rs.getString("Name"));
			row.createCell(2).setCellValue(rs.getString("Username"));
		}
		
		for (int i = 0; i < employeeHeaders.length; i++) {
			sheet_employee.autoSizeColumn(i);
		}
		rs=null;
		rs=daoFile.getResult(book_query);
		int rowNumberBook=1;
		while(rs.next())
		{
			Row row=sheet_book.createRow(rowNumberBook++);
			row.createCell(0).setCellValue(rs.getInt("BookId"));
			row.createCell(1).setCellValue(rs.getString("BookTitle"));
			row.createCell(2).setCellValue(rs.getString("Author"));
			row.createCell(3).setCellValue(rs.getString("Edition"));
		}
		
		for (int i = 0; i < bookHeaders.length; i++) {
			sheet_book.autoSizeColumn(i);
		}
		
		FileOutputStream fout=new FileOutputStream("ComLibrary.xlsx");
		workbook.write(fout);
		fout.close();
		workbook.close();
	}
}
