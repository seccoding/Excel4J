package io.github.seccoding.excel.util.write;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import io.github.seccoding.excel.util.write.share.WriteShare;

public class MakeWorkBook {

	public static void makeWorkBookAndSheet() {
		WriteShare.wb = MakeWorkBook.getWorkbook(WriteShare.writeOption.getFileName());
		WriteShare.sheet = WriteShare.wb.createSheet(WriteShare.writeOption.getSheetName());
	}

	public static Workbook getWorkbook(String fileName) {
		if (FileType.isXls(fileName)) {
			return new HSSFWorkbook();
		}
		if (FileType.isXlsx(fileName)) {
			return new SXSSFWorkbook(-1);
		}

		throw new RuntimeException("Could not find Excel file");
	}

}
