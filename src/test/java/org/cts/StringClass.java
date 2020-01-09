package org.cts;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class StringClass {
	public static void main(String[] args) throws Throwable {
		File loc = new File("C:\\Users\\Vishnu Priya\\eclipse-vishnupriya\\ExcelPractice\\Excel\\Vishnu.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("data");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				int type = c.getCellType();
				if (type == 1) {
					String name = c.getStringCellValue();
					System.out.println(name);
				}
			}

		}
	}

}
