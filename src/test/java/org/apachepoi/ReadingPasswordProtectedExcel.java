package org.apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingPasswordProtectedExcel {

	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\lenovo\\eclipse-workspace\\ApachePoiMvn\\Excel\\Customers.xlsx");
		FileInputStream stream = new FileInputStream(file);
		String password = "Test@123";
		Workbook workbook = WorkbookFactory.create(stream, password);
		// Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheetAt(0);
		int rows = sheet.getLastRowNum();
		System.out.println(rows); // 5 started from 1
		int cols = sheet.getRow(0).getLastCellNum();
		System.out.println(cols); // 3 started from 1

		for (int r = 0; r <= rows; r++) {

			Row row = sheet.getRow(r);

			for (int c = 0; c < cols; c++) {
				Cell cell = row.getCell(c);
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;

				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;

				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						Date date = cell.getDateCellValue();
						SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMMM-yy");
						System.out.print(dateFormat.format(date));

					} else {

						double d = cell.getNumericCellValue();
						BigDecimal b = BigDecimal.valueOf(d);
						System.out.print(b.toString());
						break;
					}
				case FORMULA:
					System.out.print(cell.getNumericCellValue());
					break;
				}
				System.out.print(" | ");

			}
			System.out.println();
		}
		workbook.close();
		stream.close();

	}
}
