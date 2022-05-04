package org.apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\lenovo\\eclipse-workspace\\ApachePoiMvn\\Excel\\Datas.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		// Sheet sheetAt = workbook.getSheetAt(0); we can get sheet based on index

		// USING FOR LOOP

		// int rows = sheet.getLastRowNum();
		// int cols = sheet.getRow(1).getLastCellNum();
		//
		// for (int r = 0; r < rows; r++) { // Row
		// Row row = sheet.getRow(r);
		//
		// for (int c = 0; c < cols; c++) { // Cols
		// Cell cell = row.getCell(c);
		//
		// CellType type = cell.getCellType();
		//
		// switch (type) {
		// case STRING:
		// System.out.print(cell.getStringCellValue());
		// break;
		//
		// case BOOLEAN:
		// System.out.print(cell.getBooleanCellValue());
		// break;
		//
		// case NUMERIC:
		// if (DateUtil.isCellDateFormatted(cell)) {
		// Date date = cell.getDateCellValue();
		// SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MMMM-yy");
		// System.out.print(dateFormat.format(date));
		//
		// } else {
		//
		// double d = cell.getNumericCellValue();
		// BigDecimal b = BigDecimal.valueOf(d);
		// System.out.print(b.toString());
		// break;
		// }
		//
		// default:
		// break;
		// }
		// System.out.print(" | ");
		//
		// }
		// System.out.println();
		// }

		// ************ITERATOR***************

		Iterator iterator = sheet.iterator();
		while (iterator.hasNext()) {
			Row row = (Row) iterator.next();
			
			Iterator cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = (Cell) cellIterator.next();
				
				CellType type = cell.getCellType();
				switch (type) {
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

				default:
					break;
				}
				System.out.print(" | ");

			}
			System.out.println();
		}

	}

}
