package org.cts;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Practice {
	public static void main(String[] args) throws Throwable {
		File excelloc = new File("C:\\Users\\Vishnu Priya\\eclipse-vishnupriya\\ExcelPractice\\Excel\\Vishnu.xlsx");
		FileInputStream stream = new FileInputStream(excelloc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("data");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row row = s.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				String name = "";
				if (cellType == 1) {
					name = cell.getStringCellValue();
					System.out.println(name);

				} else if (cellType == 0) {
					if (DateUtil.isCellDateFormatted(cell)) {
						Date d = cell.getDateCellValue();
						SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-YYYY");
						name = sim.format(d);
						System.out.println(name);

					} else {
						double d = cell.getNumericCellValue();
						long l = (long) d;
						name = String.valueOf(l);
						System.out.println(name);
					}

				}
			}
		}

	}

}
