package org.apachepoi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WrittingExcelDemo2 {

	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\lenovo\\eclipse-workspace\\ApachePoiMvn\\Excel\\Employee.xlsx");
		Workbook workbook = new XSSFWorkbook();
		Sheet createSheet = workbook.createSheet("Emp Info");

		ArrayList<Object[]> empdata=new ArrayList<Object[]>();
		empdata.add(new Object[]{ "EmpID", "Name", "Job" });
		empdata.add(new Object[]{ 101, "David", "Engineer" });
		empdata.add(new Object[] { 102, "Esra", "Analyst" });
		empdata.add(new Object[]{ 103, "Musa", "Manager" });
		empdata.add(new Object[] { 104, "Ozan", "Ceo" });
		
		
		
		//*********USING FOR....EACH LOOP
		
		int rowNum=0;
		for (Object[] emp : empdata) {
			Row createRow = createSheet.createRow(rowNum++);
			int columnNum=0;
			for (Object value : emp) {
				Cell createCell = createRow.createCell(columnNum++);
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
