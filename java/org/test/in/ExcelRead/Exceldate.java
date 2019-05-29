package org.test.in.ExcelRead;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exceldate {
	public static void main(String[] args) throws IOException {
		File loc = new File("C:\\Users\\ADMIN\\eclipse-workspace\\ExcelRead\\Excel\\Excelread1.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Datas");
		Date d = new Date();
		System.out.println(d);
		SimpleDateFormat Sd = new SimpleDateFormat("dd-MMM-yy");
		String name = Sd.format(d);
		System.out.println(name);
	}

}
