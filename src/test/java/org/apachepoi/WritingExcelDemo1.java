package org.apachepoi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//Workbook--->Sheet---->Rows----->Cells
public class WritingExcelDemo1 {

	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\lenovo\\eclipse-workspace\\ApachePoiMvn\\Excel\\Employee.xlsx");
		Workbook workbook = new XSSFWorkbook();
		Sheet createSheet = workbook.createSheet("Emp Info");

		Object empdata[][] = { { "EmpID", "Name", "Job" }, 
				               { 101, "David", "Engineer" }, 
				               { 102, "Esra", "Analyst" },
				               { 103, "Musa", "Manager" }, 
				               { 104, "Ozan", "Ceo" }, 
				             };
		// ***USING NORMAL FOR LOOP******

//		int rows = empdata.length;
//		int cells = empdata[0].length;
//
//		System.out.println(rows); // 5
//		System.out.println(cells); // 3
//
//		for (int i = 0; i < rows; i++) {
//			Row createRow = createSheet.createRow(i);
//
//			for (int j = 0; j < cells; j++) {
//				Cell createCell = createRow.createCell(j);
//				Object value = empdata[i][j];
//
//				if (value instanceof String)
//					createCell.setCellValue((String) value);
//				if (value instanceof Integer)
//					createCell.setCellValue((Integer) value);
//				if (value instanceof Boolean)
//					createCell.setCellValue((Boolean) value);

//			}
//		}
		
		//*********USING FOR....EACH LOOP
		
		int rowCount=0;
		for (Object emp[] : empdata) {
			Row createRow = createSheet.createRow(rowCount++);
			int columnCount=0;
			for (Object value : emp) {
				Cell createCell = createRow.createCell(columnCount++);
				if (value instanceof String)
					createCell.setCellValue((String) value);
				if (value instanceof Integer)
					createCell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					createCell.setCellValue((Boolean) value);
			}
			
		}
		
		//These things are same for 2 loop
		FileOutputStream outStream=new FileOutputStream(file);
		workbook.write(outStream);
		outStream.close();
		
		System.out.println("Employee.xlxs file written Successfully");
	}

}
