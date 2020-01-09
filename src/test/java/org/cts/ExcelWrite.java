package org.cts;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) throws Throwable {
		File loc = new File("Excel\\Priya.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("Vishnu priya");
		Row r = s.createRow(5);
		Cell c = r.createCell(5);
		c.setCellValue("Vishnu priya selected");
		FileOutputStream stream = new FileOutputStream(loc);
		w.write(stream);
		System.out.println("Wroted successfully");
	}

}
