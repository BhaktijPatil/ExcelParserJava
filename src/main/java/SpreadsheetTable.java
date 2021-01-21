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

        System.out.println("\nSpreadsheet Table created as follows :\nRows : " + rowHeaders + "\nColumns:" + colHeaders);
        new ExcelReader().displaySheet(tableData);
    }
}
