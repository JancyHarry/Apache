package org.apachepoi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell1 {

	public static void main(String[] args) throws IOException {
		
  File file =new File("C:\\Users\\lenovo\\eclipse-workspace\\ApachePoiMvn\\Excel\\Calc.xlsx");
  Workbook workbook=new XSSFWorkbook();
  
  Sheet sheet = workbook.createSheet();
  Row row = sheet.createRow(0);
  row.createCell(0).setCellValue(10);
  row.createCell(1).setCellValue(20);
  row.createCell(2).setCellValue(30);
  row.createCell(3).setCellValue(40);
  //Formula
  row.createCell(4).setCellValue("A1*B1*C1*D1");
  
  FileOutputStream outStream=new FileOutputStream(file);
  workbook.write(outStream);
  outStream.close();
  System.out.println("Calc.xlsx Created with formula cell......");
  
	}

}
