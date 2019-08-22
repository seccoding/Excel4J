package io.github.seccoding.excel.util.read;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.util.read.share.ReadShare;

public class GetSheet {

	public static void set() {
		if ( ReadShare.sheetName == null || ReadShare.sheetName.length() == 0 ) {
			GetSheet.get(0);
		}
		else {
			GetSheet.get();
		}
		setNumOfRowsAndCells();
	}
	
	public static void get(int index) {
		ReadShare.sheet = ReadShare.wb.getSheetAt(index);
	}
	
	public static void get() {
		ReadShare.sheet = ReadShare.wb.getSheet(ReadShare.sheetName);
	}
	
	public static void getSheetName() {
		
		if ( ReadShare.readOption.getSheetName() == null ) {
			ExcelSheet sheet = ReadShare.clazz.getAnnotation(ExcelSheet.class);
			if ( sheet != null ) {
				ReadShare.sheetName = sheet.value();
			}
			set();
		}
	}
	
	public static void setNumOfRowsAndCells() {
		if ( ReadShare.sheet == null ) {
			throw new RuntimeException("Can not find sheet [" + ReadShare.sheetName + "]");
		}
		ReadShare.numOfRows = ReadShare.sheet.getPhysicalNumberOfRows();
		ReadShare.numOfCells = 0;
	}
	
}
