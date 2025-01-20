package org.example;

import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        // Specify the path of the Excel file
        String excelFilePath = "/C:/Users/sarth/Downloads/DemoExcel.xlsx~"; // Update the file path

        // Reading the Excel file
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath))) {
            // Create a Workbook instance to read the Excel file
            Workbook workbook = WorkbookFactory.create(fis);

            // Get the first sheet (you can change the index to get other sheets)
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through rows
            for (Row row : sheet) {
                // Iterate through cells of each row
                for (Cell cell : row) {
                    // Based on the cell type, read the appropriate value
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            // Check if it's a date or a number
                            if (DateUtil.isCellDateFormatted(cell)) {
                                System.out.print(cell.getDateCellValue() + "\t");
                            } else {
                                System.out.print(cell.getNumericCellValue() + "\t");
                            }
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        default:
                            System.out.print("N/A\t");
                    }
                }
                System.out.println(); // Move to the next row
            }

            // Close the workbook after use
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace(); // Print exception stack trace if any error occurs
        }
    }
}
