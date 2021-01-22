import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.util.HashMap;


public class ExcelParser {

    private static ExcelReader excelReader = new ExcelReader();
    private static TableParser tableParser = new TableParser();
    private static int headerRowIndex = 2;
    private static int headerColumnIndex = 1;
    private static HashMap errorMap = new HashMap();
    private static String fileLocation = "exceldata/Financial Sample.xlsx";
//    private static String fileLocation = "exceldata/Sample Excel.xlsx";

    public static void main(String[] args) {
        System.out.println("Application Started");
        Sheet sheet;
        try {
            // Read Excel sheet
//             sheet = excelReader.getSheetAtIndex(fileLocation, 0, ExcelReader.XLSX);
            sheet = excelReader.getSheetByName(fileLocation, "Formatted", ExcelReader.XLSX);

            SpreadsheetTable tableData = tableParser.parse(sheet, headerRowIndex, headerColumnIndex);
            tableData.displayTable();

            int rowNum = 5;
            int colNum = 1;
            Cell cell = excelReader.fetchCell(rowNum, colNum, tableData.tableData);
            excelReader.displayCellContent(rowNum, colNum, tableData.tableData);
            tableData.displayCellContent("Sale Price", "Canada");

        } catch (IOException ioException) {
            System.out.println("Excel Sheet not found at given location. ERROR : " + ioException);
        }
    }

    // Function to validate datatype in a column()
    private static HashMap validateIntegerRow(Row row, HashMap errorMap) {
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.NUMERIC) {
                errorMap.put(new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(), "\tFound: " + cell.getCellType() + "\tExpected: " + CellType.NUMERIC);
            } else if (DateUtil.isCellDateFormatted(cell))
                errorMap.put(new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(), "Found: Date\tExpected: " + CellType.NUMERIC);
        }
        return errorMap;
    }

    // Function to apply rule to a column()
    private static HashMap filterRowManuPrice300(Sheet sheet, HashMap errorMap) {
        for (Row row : sheet) {
            Cell cell = row.getCell(5);
            if (cell.getCellType() == CellType.NUMERIC)
                if (cell.getNumericCellValue() > 300)
                    errorMap.put(new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(), "Found: " + cell.getNumericCellValue() + "\tExpected: <300");
        }
        return errorMap;
    }

    //        try {
//
//            printSheet(sheet);
//
//            // Find validation errors
//            errorMap.putAll(validateIntegerRow(sheet.getRow(5), errorMap));
//            errorMap.putAll(filterRowManuPrice300(sheet, errorMap));
//
//            System.out.println("\n\nERROR LOG :");
//            System.out.println(errorMap);
//        } catch (Exception e) {
//            System.out.println("Excel Sheet not found at given location " + e);
//        }

}
