import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

public class SpreadsheetTable {
    public ArrayList<String> rowHeaders;
    public ArrayList<String> colHeaders;
    public Sheet tableData;

    public SpreadsheetTable(ArrayList<String> rowHeaders, ArrayList<String> colHeaders, Sheet tableData) {
        this.rowHeaders = rowHeaders;
        this.colHeaders = colHeaders;
        this.tableData = tableData;
    }

    // Function to print table details
    public void displayTable() {
        System.out.print("\nSpreadsheet Table Details :\nRows : " + rowHeaders + "\nColumns:" + colHeaders);
        new ExcelReader().displaySheet(tableData);
    }

    // Function to fetch a cell from table given row and column names
    public Cell fetchCell(String rowName, String colName) {
        int rowNum = rowHeaders.indexOf(rowName);
        int colNum = rowHeaders.indexOf(colName);

        return new ExcelReader().fetchCell(rowNum, colNum, tableData);
    }

    // Function to display cell content given row and column names
    public void displayCellContent(String rowName, String colName) {
        int rowNum = rowHeaders.indexOf(rowName);
        int colNum = colHeaders.indexOf(colName);

        new ExcelReader().displayCellContent(rowNum, colNum, tableData);
    }
}
