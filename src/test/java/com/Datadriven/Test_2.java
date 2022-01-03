package com.Datadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test_2 {
	public static void main(String[] args) throws Throwable {
		
		File f = new File("C:\\Users\\nivim\\eclipse-workspace\\Datadriven\\Datadriven.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		
		Sheet createSheet = wb.createSheet("Data");
		Row createRow = createSheet.createRow(0);
		Cell createCell = createRow.createCell(0);
		
		createCell.setCellValue("User Data");
		wb.getSheet("Data").getRow(0).createCell(1).setCellValue("User Name");
		wb.getSheet("Data").createRow(1).createCell(0).setCellValue("Gowdha");
		wb.getSheet("Data").createRow(1).createCell(1).setCellValue("Tom");
		
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		System.out.println("Data sheet created Succesfully");

				
			}
		}
	
			
		
