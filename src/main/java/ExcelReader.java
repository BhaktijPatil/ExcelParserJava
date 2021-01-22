//import org.apache.poi.sl.usermodel.Sheet;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {

    public static final int XLS = 0;
    public static final int XLSX = 1;

    // Function to read an excel sheet given the file location and sheet index
    public Sheet getSheetAtIndex(String excelFileLoc, int index, int excelFormat) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(excelFileLoc);
        Workbook workbook = excelFormat == XLS ? new HSSFWorkbook(fileInputStream) : new XSSFWorkbook(fileInputStream);
        // Get sheet by sheet index
        return workbook.getSheetAt(index);
    }

    // Function to read an excel sheet given the file location and sheet name
    public Sheet getSheetByName(String excelFileLoc, String sheetName, int excelFormat) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(excelFileLoc);
        Workbook workbook = excelFormat == XLS ? new HSSFWorkbook(fileInputStream) : new XSSFWorkbook(fileInputStream);
        // Get sheet by sheet index
        return workbook.getSheet(sheetName);
    }

    // Function to print sheet
    public void displaySheet(Sheet sheet, int headerRowIndex) {
        System.out.println("\n\nSheet Name : " + sheet.getSheetName());
        // Print column headers
        System.out.print("\t");
        for (Cell cell : sheet.getRow(headerRowIndex))
            System.out.print(String.format("%-25s", CellReference.convertNumToColString(cell.getColumnIndex())));
        // Print rows
        for (Row row : sheet) {
            System.out.print("\n" + row.getRowNum() + "\t");
            for (Cell cell : row) {
                displayCellValue(cell);
            }
        }
        System.out.println("\n");
    }

    // Function to print Cell w/ proper formatting
    private void displayCellValue(Cell cell) {
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

    // Function to print sheet
    public void displaySheet(Sheet sheet) {
        displaySheet(sheet, sheet.getTopRow());
    }

    // Function to display contents of a cell
    public void displayCellContent(int rowNum, int colNum, Sheet sheet) {
        Cell cell = fetchCell(rowNum, colNum, sheet);
        System.out.print("\nCell value at :\nRow : " + rowNum + "  Col : " + colNum + "  Value : ");
        displayCellValue(cell);
        System.out.println("\n");
    }

    // Function to fetch a cell from table given row and column IDs
    public Cell fetchCell(int rowNum, int colNum, Sheet sheet) {
        return sheet.getRow(rowNum).getCell(colNum);
    }
}
