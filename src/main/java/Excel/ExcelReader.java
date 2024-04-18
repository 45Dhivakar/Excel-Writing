package Excel;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    public static void main(String[] args) {
        try {
            // Open the Excel file
            FileInputStream file = new FileInputStream("output.xlsx");

            // Create a workbook instance
            Workbook workbook = new XSSFWorkbook(file);

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows
            for (Row row : sheet) {
                // Iterate through cells
                for (Cell cell : row) {
                    // Check the cell type and print the value accordingly
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        case BLANK:
                            System.out.print("[BLANK]\t");
                            break;
                        default:
                            System.out.print("[UNKNOWN]\t");
                    }
                }
                System.out.println(); // Move to the next line after each row
            }

            // Close the workbook and file
            workbook.close();
            file.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}



