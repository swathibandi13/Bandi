package org.test.in.ExcelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import com.google.common.collect.Table.Cell;

public class PexcelData {
	public static void main(String[] args) throws IOException {
		// file loc

		File loc = new File("C:\\Users\\ADMIN\\eclipse-workspace\\ExcelRead\\Excel\\Excelread1.xlsx");

		// converting object

		FileInputStream stream = new FileInputStream(loc);

		// workbook

		Workbook w = new XSSFWorkbook(stream);

		// sheet

		Sheet s = w.getSheet("Datas");

		// row

		/*
		 * Row r = s.getRow(7);
		 * 
		 * // cell
		 * 
		 * Cell d = r.getCell(1);
		 * 
		 * System.out.println(d);
		 */

		// print no . of rows

		/*
		 * int count = s.getPhysicalNumberOfRows(); int count1
		 * =r.getPhysicalNumberOfCells();
		 */

		// System.out.println(count);
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {

			Row r = s.getRow(i);

			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell d = r.getCell(j);

				int type = d.getCellType();

				if (type == 1) {
					String name = d.getStringCellValue();
					System.out.println(name);

				}

				if (type == 0) {
					double d1 = d.getNumericCellValue();
					// double -long

					long l = (long) d1;
					// long-string

					String name = String.valueOf(l);
					System.out.println(name);

				}
			}

		}
	}
}
