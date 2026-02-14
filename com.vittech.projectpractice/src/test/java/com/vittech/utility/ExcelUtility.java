    package com.vittech.utility;

	import java.io.FileInputStream;
	import java.io.IOException;

	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class ExcelUtility {

	    private Workbook workbook;
	    private Sheet sheet;

	    // Constructor
	    public ExcelUtility(String filePath, String sheetName) throws IOException {
	        FileInputStream fis = new FileInputStream(filePath);
	        workbook = new XSSFWorkbook(fis);
	        sheet = workbook.getSheet(sheetName);
	    }

	    // Get row count
	    public int getRowCount() {
	        return sheet.getPhysicalNumberOfRows();
	    }

	    // Get column count
	    public int getColumnCount() {
	        return sheet.getRow(0).getPhysicalNumberOfCells();
	    }

	    // Get single cell data
	    public String getCellData(int rowNum, int colNum) {
	        Cell cell = sheet.getRow(rowNum).getCell(colNum);
	        return cell.toString();
	    }

	    // Get all data (2D Object array) 
	    public Object[][] getSheetData() {

	        int rows = getRowCount();
	        int cols = getColumnCount();

	        Object[][] data = new Object[rows - 1][cols];

	        for (int i = 1; i < rows; i++) {
	            for (int j = 0; j < cols; j++) {
	                data[i - 1][j] = getCellData(i, j);
	            }
	        }

	        return data;
	    }

	    public void close() throws IOException {
	        workbook.close();
	    }
	}

