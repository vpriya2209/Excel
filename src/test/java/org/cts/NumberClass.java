package org.cts;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NumberClass {
	public static void main(String[] args) throws Throwable {
		File excellocc = new File("Excel\\Vishnu.xlsx");
		FileInputStream stream = new FileInputStream(excellocc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("data");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				int type = c.getCellType();
				if (type == 0) {
					double d = c.getNumericCellValue();
					long l = (long) d;
					String name = String.valueOf(l);
					System.out.println(name);
				}

			}

		}
	}

}
