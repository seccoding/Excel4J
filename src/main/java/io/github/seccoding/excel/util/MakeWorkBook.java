package io.github.seccoding.excel.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MakeWorkBook {

	public static Workbook getWorkbook(String fileName) {

		if (FileType.isXls(fileName)) {
			return new HSSFWorkbook();
		}
		if (FileType.isXlsx(fileName)) {
			return new XSSFWorkbook();
		}
		
		throw new RuntimeException("Could not find Excel file");

	}

}
