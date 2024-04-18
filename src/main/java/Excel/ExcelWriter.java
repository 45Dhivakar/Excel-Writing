package Excel;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	public static void main(String[] args) {
		try {
			// Create a new Excel Workbook
			Workbook workbook = new XSSFWorkbook();
			
			// Crate a new sheet with the name "Sheet1"
			Sheet sheet = workbook.createSheet("Sheet1");
			// Write column headers 
			Row headerRow =  sheet.createRow(0);
			String[] headers = {"Name", "Age", "Email"};
			for (int i =0; i < headers.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(headers[i]);
			}
			
			// Write data rows
			String[][] data = {
					{"Ram", "30", "ram@test.com"},
					{"Raja", "40", "raja@test.com"},
					{"sri", "26", "Sri@test.com"},
					{"Jai","18",  "jai@test.com"}
			};
			
			for (int i = 0; i< data.length; i++) {
				Row row = sheet.createRow(i+1);
				for (int j = 0; j< data[i].length; j++) {
					Cell cell = row.createCell(j);
					cell.setCellValue(data[i][j]);
					
				}
			}
			// Write the workbook to a file
			FileOutputStream fileOut = new FileOutputStream("output.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			
			// Close the workbook
			workbook.close();
			
			System.out.println("Excel file has been created sucessfully!");
		}catch (Exception e) {
			e.printStackTrace();
		}
}
}