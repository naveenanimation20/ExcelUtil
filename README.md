# ExcelUtil
An excel utility for Excel read and write operations.

This is a Java class that allows reading data from an Excel (.xlsx) file. It provides methods for getting the row count and cell data of a specific sheet.

The class uses the Apache POI library for working with Excel files. It has a constructor that takes the path of the Excel file as a parameter and initializes the workbook object.

The class has two methods for getting cell data, one that takes the sheet name and column name as parameters, and another that takes the sheet name and column number as parameters. Both methods also require the row number as a parameter.

The class uses the Apache POI library's cell object to get the data from the cell. It handles different types of cell data such as string, numeric, formula, and date.

The class also has a method for getting the row count of a specific sheet. It takes the sheet name as a parameter and uses the workbook object to get the sheet index and then returns the number of rows in the sheet.


Here is a more detailed description of each method in the provided code:

**public NALExcelXLSReader(String path)**: 
This is the constructor of the NALExcelXLSReader class. 
It takes a file path as input and initializes the instance variables path, fis, workbook, sheet, row, and cell.
It also reads the workbook from the given file path using a FileInputStream and initializes the workbook and sheet instance variables to the first sheet of the workbook.

**public int getRowCount(String sheetName)**: 
This method takes a sheet name as input and returns the number of rows in the sheet. 
It first gets the index of the sheet using the workbook.getSheetIndex(sheetName) method. 
If the sheet doesn't exist, it returns 0. Otherwise, it gets the sheet using workbook.getSheetAt(index) and returns the number of rows in the sheet using sheet.getLastRowNum() + 1.

**public String getCellData(String sheetName, String colName, int rowNum)**: 
This method takes a sheet name, column name, and row number as input and returns the data in the cell corresponding to the input parameters. 
It first gets the index of the sheet using the workbook.getSheetIndex(sheetName) method. 
If the sheet doesn't exist, it returns an empty string. 
Otherwise, it searches for the column number corresponding to the input column name by iterating over the first row of the sheet. 
If the column doesn't exist, it returns an empty string. Otherwise, it gets the cell using row.getCell(col_Num) and returns the cell data as a string.
If the cell is a string, it returns the string value. 
If the cell is a number or formula, it returns the numeric value as a string. 
If the cell is a blank, it returns an empty string. 
If the cell is a boolean, it returns the boolean value as a string.

**public String getCellData(String sheetName, int colNum, int rowNum)**:
This method takes a sheet name, column number, and row number as input and returns the data in the cell corresponding to the input parameters.
It first gets the index of the sheet using the workbook.getSheetIndex(sheetName) method.
If the sheet doesn't exist, it returns an empty string. Otherwise, it gets the cell using row.getCell(colNum) and returns the cell data as a string.
If the cell is a string, it returns the string value. 
If the cell is a number or formula, it returns the numeric value as a string. 
If the cell is a blank, it returns an empty string. If the cell is a boolean, it returns the boolean value as a string.

Note that both getCellData methods throw an exception and return an error message if the row or column doesn't exist in the sheet.
