package org.cts;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DateFormateClass {
	public static void main(String[] args) throws Throwable {
		File loc = new File("Excel\\Vishnu.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("data");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell cell = r.getCell(j);
				int cellType = cell.getCellType();
				if (cellType == 0) {
					if (DateUtil.isCellDateFormatted(cell));
					java.util.Date date = cell.getDateCellValue();
					SimpleDateFormat sim = new SimpleDateFormat("dd-MM-YYYY");
					String format = sim.format(date);
					System.out.println(format);

				}

			}

		}
	}
}
