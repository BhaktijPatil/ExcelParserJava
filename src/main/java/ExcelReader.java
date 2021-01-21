//import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

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
}

//    // Function to read an excel sheet given the file location and sheet index
//    public Sheet getSheetAtIndex(String excelFileLoc, int index, int excelFormat) throws IOException {
//        FileInputStream fileInputStream = new FileInputStream(excelFileLoc);
//        // Read excel file
//        XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
//        // Get sheet by sheet index
//        return wb.getSheetAt(index);
//    }
