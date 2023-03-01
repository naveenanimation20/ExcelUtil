package com.tests;

import java.util.Arrays;

import com.navlabs.excel.reader.NALExcelXLSReader;

public class ExcelPOITest {

	public static void main(String[] args) {

		
		NALExcelXLSReader reader = new NALExcelXLSReader("testdata.xlsx");
		int col = reader.getColumnCount("register");
		System.out.println(col);
		
		String cell = reader.getCellData("register", "firstname", 2);
		System.out.println(cell);
		
		reader.addSheet("naveen");
		
		Object data[][] = reader.getSheetData("register");
		System.out.println(Arrays.deepToString(data));
	}

}
