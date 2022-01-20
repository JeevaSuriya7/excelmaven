package org.exe;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exec {
public static void main(String[] args) throws Exception {
	File file = new File("C:\\Users\\Admin\\eclipse-workspace\\Excel\\Exxcel\\empDetails.xlsx");
	FileInputStream stream = new FileInputStream(file);
	Workbook w =new XSSFWorkbook(stream);
	Sheet sheet = w.getSheet("Sheet1");
	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
		Row row = sheet.getRow(i);
		for (int j = 0; j <row.getPhysicalNumberOfCells(); j++) {
			Cell cell = row.getCell(j);
			System.out.println(cell);
		}
		
	}
	
	
	
	
}
}
