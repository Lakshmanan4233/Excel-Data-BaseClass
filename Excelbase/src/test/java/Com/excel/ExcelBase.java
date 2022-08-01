package Com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelBase {
	
	
	// Print one data in console from excel sheet
		public void sheet1(String pathOfFile,String sheetName,int rowNo,int cellNo) throws IOException {
			
			File file = new File(pathOfFile);
			
			FileInputStream fileInputStream  = new FileInputStream(file);
		
			Workbook workbook = new XSSFWorkbook(fileInputStream);
			
			Sheet sheet = workbook.getSheet(sheetName);
			
			Row row = sheet.getRow(rowNo);
			
			Cell cell = row.getCell(cellNo);
			
			System.out.println(cell);
			
		}
		
		
		//get no of row in sheet
		public void noOfRows(String path,String sheetName) throws Exception {
			
			File file = new File(path);
			
			FileInputStream fileInputStream = new FileInputStream(file);
			
			Workbook workbook = new XSSFWorkbook(fileInputStream);
			
			Sheet sheet = workbook.getSheet(sheetName);
			
			System.out.println(sheet.getPhysicalNumberOfRows());
			
		}
		
		//get no of cell in sheet
		
		public void noOfCell(String path,String sheetName,int rowNo) throws IOException {
			
			File file = new File(path);
			
			FileInputStream fileInputStream = new FileInputStream(file);
			
			Workbook workbook = new XSSFWorkbook(fileInputStream);
			
			Sheet sheet = workbook.getSheet(sheetName);
			
			Row row = sheet.getRow(rowNo);
			
			System.out.println(row.getPhysicalNumberOfCells());

		}
		
		//get all cell data in particular row
		public void getCellData(String path,String sheetName,int rowNo) throws Exception {
			
			File file = new File(path);
			
			FileInputStream fileInputStream = new FileInputStream(file);
			
			Workbook workbook = new XSSFWorkbook(fileInputStream);
			
			Sheet sheet = workbook.getSheet(sheetName);
			
			Row row = sheet.getRow(rowNo);
			
			
			for(int i=0; i<row.getPhysicalNumberOfCells();i++) {
				
				System.out.println(row.getCell(i));
			}

		}
		
		
		//get the all data in excel
		
		public void getAllData(String path,String sheetName) throws IOException {
			
			File file = new File(path);
			
			FileInputStream fileInputStream = new FileInputStream(file);
			
			Workbook workbook = new XSSFWorkbook(fileInputStream);
			
			Sheet sheet = workbook.getSheet(sheetName);
			
			for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
				
				Row row = sheet.getRow(i);
				
				for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
					
					Cell cell = row.getCell(j);
					
					System.out.println(cell);
				}
			}
			
			

		}
		

}
