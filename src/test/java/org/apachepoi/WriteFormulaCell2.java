package org.apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell2 {

	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\lenovo\\eclipse-workspace\\ApachePoiMvn\\Excel\\Books.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");
		stream.close();
		FileOutputStream outStream = new FileOutputStream(file);
		workbook.write(outStream);
		outStream.close();
		System.out.println("formula cell is Created......");

	}

}
