package com.Datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test_1 {
	public static void main(String[] args) throws Throwable {
		File f = new File("C:\\Users\\nivim\\eclipse-workspace\\Datadriven\\Datadriven.xlsx");
		
		FileInputStream fis = new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet sheetAt = wb.getSheetAt(0);
		int rows = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {
			Row row2 = sheetAt.getRow(i);
			int call = row2.getPhysicalNumberOfCells();
			for (int j = 0; j < call; j++) {
				Cell cell = row2.getCell(j);
				CellType cellType = cell.getCellType();
				
				if (cellType.equals(cellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
					
				} else if (cellType.equals(cellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int a = (int) numericCellValue;
					System.out.println(a);
					
				}
				
				
				
				
				
			}
			
			
			
		}
		
		
		
	}

	}


