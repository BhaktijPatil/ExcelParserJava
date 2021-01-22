# ExcelParserJava
Maven project to Parse data from Excel files using .xlsx/.xls format. Project allows to parse Excel files with ease using Apache POI, and also includes parsers for Tables inside Excel sheets.

# Screenshots
### READ SHEET FROM EXCEL FILE :
![Reading Excel Sheet](https://github.com/BhaktijPatil/ExcelParserJava/blob/master/results/excel_sheet_print.png)

### READ TABULAR DATA FROM EXCEL FILE :
![Reading Table](https://github.com/BhaktijPatil/ExcelParserJava/blob/master/results/formatted_%20table.png)

### FETCH DATA FROM PARSED TABLE :
![Fetch Cell](https://github.com/BhaktijPatil/ExcelParserJava/blob/master/results/fetch_cell_value.png)

# Documentation :
## SheetParser :
### Constructors :
    // Read an excel sheet given the file location and sheet index
    public SheetParser(String excelFileLoc, int index, int excelFormat) 
    // Read an excel sheet given the file location and sheet name
    public SheetParser(String excelFileLoc, String sheetName, int excelFormat)
    // Constructor to initialize a sheet directly
    public SheetParser(Sheet sheet)
### Public Functions :
    // Function to print sheet
    public void displaySheet(int headerRowIndex)
    // Function to print sheet
    public void displaySheet()
    // Function to print Cell Value w/ proper formatting
    private void displayCellValue(Cell cell)
    // Function to display contents of a cell
    public void displayCellContent(int rowNum, int colNum)
    // Function to fetch a cell from table given row and column IDs
    public Cell fetchCell(int rowNum, int colNum)
    
## SpreadsheetTable :
### Constructors :
    // Constructor to load table from a sheet given the row headings, column headings & sheet
    public SpreadsheetTable(ArrayList<String> rowHeaders, ArrayList<String> colHeaders, Sheet tableData)
### Public Functions :
    // Function to print table details
    public void displayTable()
    // Function to fetch a cell from table given row and column names
    public Cell fetchCell(String rowName, String colName)
    // Function to display cell content given row and column names
    public void displayCellContent(String rowName, String colName)

## TableParser :
### Public Functions:
    // Parses table within the specified bounds
    public SpreadsheetTable parse(Sheet sheet, int headerRowIndex, int headerColIndex, int lastRowIndex, int lastColIndex)
    // Parses Table till first Blank Row & Column are found
    public SpreadsheetTable parse(Sheet sheet, int headerRowIndex, int headerColIndex) 
    // Function to return Column Names
    private ArrayList<String> getColHeaders(Sheet sheet, int headerColindex)
    // Function to return Row Names
    private ArrayList<String> getRowHeaders(Sheet sheet, int headerRowIndex)
