package Excel;

	import java.io.FileOutputStream;
	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class ExcelWriter1 {
	    public static void main(String[] args) {
	        // Create a new Excel workbook
	        Workbook workbook = new XSSFWorkbook();

	        // Create a new sheet with the name "Sheet1"
	        Sheet sheet = workbook.createSheet("Sheet1");

	        // Write column headers
	        Row headerRow = sheet.createRow(0);
	        String[] headers = {"Name", "Age", "Email"};
	        for (int i = 0; i < headers.length; i++) {
	            Cell cell = headerRow.createCell(i);
	            cell.setCellValue(headers[i]);
	        }

	        // Write data rows
	        String[][] data = {
	            {"John Doe", "30", "john@test.com"},
	            {"Jane Doe", "28", "jane@test.com"},
	            {"Bob Smith", "35", "bob@example.com"},
	            {"Swapnil", "37", "swapnil@example.com"}
	        };

	        for (int i = 0; i < data.length; i++) {
	            Row row = sheet.createRow(i + 1);
	            for (int j = 0; j < data[i].length; j++) {
	                Cell cell = row.createCell(j);
	                cell.setCellValue(data[i][j]);
	            }
	        }

	        // Write the workbook to a file
	        try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
	            workbook.write(fileOut);
	            System.out.println("Excel file has been created successfully!");
	        } catch (Exception e) {
	            e.printStackTrace();
	        } finally {
	            // Close the workbook
	            try {
	                workbook.close();
	            } catch (Exception e) {
	                e.printStackTrace();
	            }
	        }
	    }
	}



