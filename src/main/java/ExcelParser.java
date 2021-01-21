import org.apache.commons.logging.Log;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Formatter;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;


public class ExcelParser {
    public static void main(String[] args) {
        System.out.println("Application Started");

        // Read Excel sheet
        String fileLocation = "exceldata/Financial Sample.xlsx";
        //String fileLocation = "exceldata/sampleexcel.xlsx";

        ExcelReader excelReader = new ExcelReader();
        HashMap errorMap = new HashMap();

        try {
            Sheet sheet = excelReader.getSheetAtIndex(fileLocation, 0, ExcelReader.XLSX);
            printSheet(sheet);

            // Find validation errors
            errorMap.putAll(validateIntegerRow(sheet.getRow(5), errorMap));
            errorMap.putAll(filterRowManuPrice300(sheet, errorMap));

            System.out.println("\n\nERROR LOG :");
            System.out.println(errorMap);
        } catch (Exception e) {
            System.out.println("Excel Sheet not found at given location " + e);
        }
    }

    // Function to validate datatype in a column()
    private static HashMap validateIntegerRow(Row row, HashMap errorMap) {
        for (Cell cell : row) {
            if(cell.getCellType() != CellType.NUMERIC)
            {
                errorMap.put(new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(), "\tFound: " + cell.getCellType() + "\tExpected: " + CellType.NUMERIC);
            }
            else if(DateUtil.isCellDateFormatted(cell))
                errorMap.put(new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(), "Found: Date\tExpected: " + CellType.NUMERIC);
        }
        return errorMap;
    }

    // Function to apply rule to a column()
    private static HashMap filterRowManuPrice300(Sheet sheet, HashMap errorMap) {
        for (Row row : sheet) {
            Cell cell = row.getCell(5);
            if(cell.getCellType() == CellType.NUMERIC)
                if(cell.getNumericCellValue() > 300)
                    errorMap.put(new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(), "Found: " + cell.getNumericCellValue() + "\tExpected: <300");
        }
        return errorMap;
    }



    // Function to print sheet
    private static void printSheet(Sheet sheet) {
        // Print column headers
        System.out.print("\t");
        for (Cell cell : sheet.getRow(0))
            System.out.print(String.format("%-25s", cell.getColumnIndex()));
        // Print rows
        for (Row row : sheet) {
            System.out.print("\n" + row.getRowNum() + "\t");
            for (Cell cell : row) {
                // Format data before printing
                switch (cell.getCellType()) {
                    case STRING:
                        String cellValue = cell.getRichStringCellValue().getString();
                        if (cellValue.length() > 20)
                            cellValue = cellValue.substring(0, 20) + "..";
                        System.out.print(String.format("%-25s", cellValue));
                        break;
                    case _NONE:
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            System.out.print(String.format("%-25s", cell.getDateCellValue().toString().substring(0, 20) + ".."));
                        } else {
                            System.out.print(String.format("%-25s", cell.getNumericCellValue()));
                        }
                        break;
                    case BOOLEAN:
                        System.out.print(String.format("%-25s", cell.getBooleanCellValue()));
                        break;
                    case FORMULA:
                        System.out.print(String.format("%-25s", cell.getCellFormula()));
                        break;
                    case ERROR:
                        System.out.print("ERROR");
                        break;
                    default:
                        System.out.print("");
                }
            }
        }
    }

}
