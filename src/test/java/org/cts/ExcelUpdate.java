package org.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdate {
public static void main(String[] args) throws Throwable {
	File loc = new File("Excel\\Vishnu.xlsx"); 
	FileInputStream stream = new FileInputStream(loc);
	Workbook w = new XSSFWorkbook(stream);
	Sheet S = w.getSheet("data");
	Row r = S.getRow(0);
	Cell c = r.getCell(0);
	String name = c.getStringCellValue();
	if (name.contains("Emp")) {
		c.setCellValue("Name");
	}
	FileOutputStream o = new FileOutputStream(loc);
	w.write(o);
	System.out.println("Done");
	System.out.println(" Updated suceesfully");
}
}
