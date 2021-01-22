import org.apache.poi.ss.usermodel.*;

import java.io.IOException;


public class ExcelParser {

    private static SheetParser sheetParser;
    private static TableParser tableParser = new TableParser();
    private static int headerRowIndex = 2;
    private static int headerColumnIndex = 1;
    private static String fileLocation = "exceldata/Financial Sample.xlsx";

    public static void main(String[] args) {
        System.out.println("Application Started");
        Sheet sheet;
        try {
            // Read Excel sheet
            sheetParser = new SheetParser(fileLocation, "Formatted", SheetParser.XLSX);
            // sheetParser = new SheetParser(fileLocation, "Formatted", SheetParser.XLSX);

            SpreadsheetTable tableData = tableParser.parse(sheetParser.sheet, headerRowIndex, headerColumnIndex);
            tableData.displayTable();

            int rowNum = 5;
            int colNum = 1;
            Cell cell = sheetParser.fetchCell(rowNum, colNum);
            sheetParser.displayCellContent(rowNum, colNum);
            tableData.displayCellContent("Sale Price", "Canada");

        } catch (IOException ioException) {
            System.out.println("Excel Sheet not found at given location. ERROR : " + ioException);
        }
    }

}
