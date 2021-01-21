import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

public class TableParser {

    // Parses table within the specified bounds
    public SpreadsheetTable parse(Sheet sheet, int headerRowIndex, int headerColIndex, int lastRowIndex, int lastColIndex) {
        System.out.println("\nTable Details :\nSheet name : " + sheet.getSheetName() + "\nHeader Row Index : " + headerRowIndex + "\nHeader Col Index : " + headerColIndex + "\nLast Row Index : " + lastRowIndex + "\nLast Col Index : " + lastColIndex);

        // Clone sheet
        Sheet tableData = sheet.getWorkbook().cloneSheet(sheet.getWorkbook().getSheetIndex(sheet.getSheetName()));

        // Remove unnecessary rows and columns from table
        for (int rowNum = lastRowIndex; rowNum <= tableData.getLastRowNum(); rowNum++) {
            if (tableData.getRow(rowNum) != null) {
                tableData.removeRow(tableData.getRow(rowNum));
            }
        }
        tableData.shiftRows(headerRowIndex, lastRowIndex, -headerRowIndex);
        for (Row row : tableData)
        {
            for (int colNum = lastColIndex; colNum <= row.getLastCellNum(); colNum++) {
                if (row.getCell(colNum) != null) {
                    System.out.println(row.getRowNum() + "\t" + colNum);
                    row.removeCell(row.getCell(colNum));
                }
            }
        }
        tableData.shiftColumns(headerColIndex, lastColIndex, -headerColIndex);

        // Store row and column headers in ArrayLists
        ArrayList<String> rowHeaders = getRowHeaders(tableData, 0);
        ArrayList<String> colHeaders = getColHeaders(tableData, 0);

        // Return formatted table
        return new SpreadsheetTable(rowHeaders, colHeaders, tableData);
    }

    // Parses Table till first Blank Row & Column are found
    public SpreadsheetTable parse(Sheet sheet, int headerRowIndex, int headerColIndex) {
        int lastRowIndex = headerRowIndex;
        int lastColIndex = headerColIndex;
        for (; sheet.getRow(lastRowIndex) != null && sheet.getRow(lastRowIndex).getCell(headerColIndex).getCellType() != CellType.BLANK && lastRowIndex < sheet.getLastRowNum(); lastRowIndex++)
            ;
        for (; sheet.getRow(headerRowIndex).getCell(lastColIndex).getCellType() != CellType.BLANK && lastColIndex < sheet.getRow(headerRowIndex).getLastCellNum(); lastColIndex++)
            ;
        return parse(sheet, headerRowIndex, headerColIndex, lastRowIndex, lastColIndex);
    }

    // Function to return Column Names
    private ArrayList<String> getColHeaders(Sheet sheet, int headerColindex)
    {
        ArrayList<String> colHeaders = new ArrayList<>();
        sheet.getRow(headerColindex).forEach(cell -> {colHeaders.add(cell.getStringCellValue());});
        return colHeaders;
    }

    // Function to return Row Names
    private ArrayList<String> getRowHeaders(Sheet sheet, int headerRowIndex)
    {
        ArrayList<String> rowHeaders = new ArrayList<>();
        for(Row row : sheet)
            rowHeaders.add(row.getCell(headerRowIndex).getStringCellValue());
        return rowHeaders;
    }
}
