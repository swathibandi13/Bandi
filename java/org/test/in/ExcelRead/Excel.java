package org.test.in.ExcelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws IOException {
		// file loc

		File loc = new File("C:\\Users\\ADMIN\\eclipse-workspace\\ExcelRead\\Excel\\Excelread1.xlsx");

		// converting object

		FileInputStream stream = new FileInputStream(loc);

		// workbook

		Workbook w = new XSSFWorkbook(stream);

		// sheet

		Sheet s = w.getSheet("Datas");

		// System.out.println(count);
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {

			Row r = s.getRow(i);

			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				System.out.println(c);
			
				// 1=text , 0 = number
				int type = c.getCellType();

				if (type == 1) {
					String name = c.getStringCellValue();
					System.out.println(name);

				}

				if (type == 0) {
					double d = c.getNumericCellValue();
					System.out.println(d);
					// double -long

					long l = (long)d;
					// long-string

					String name = String.valueOf(l);
					System.out.println(name);

				}

			}

		}
	}
}
