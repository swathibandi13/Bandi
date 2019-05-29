package org.test.in.ExcelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.formula.functions.Address;
//import org.apache.poi.ss.formula.functions.Address;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateExcel {
	public static void main(String[] args) throws IOException {
		File loc = new File("C:\\Users\\ADMIN\\eclipse-workspace\\ExcelRead\\Excel\\Excelread1.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Datas");
		Row r = s.getRow(3);
		Cell c = r.getCell(2);
        String address = c.getStringCellValue();
		if (address.equals("adyar")) {
			c.setCellValue("perungudi");

		}
		// conv File
		FileOutputStream fo = new FileOutputStream(loc);
		// write in work book
		w.write(fo);
		System.out.println("Done...");

	}
}