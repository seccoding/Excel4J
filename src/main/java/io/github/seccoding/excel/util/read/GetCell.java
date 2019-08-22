package io.github.seccoding.excel.util.read;

import io.github.seccoding.excel.util.read.share.ReadShare;

public class GetCell {

	public static void setCell(int cellIndex) {
		ReadShare.cell = ReadShare.row.getCell(cellIndex);
	}
	
	public static String getCellName(int cellIndex) {
		return CellReferenceUtil.getName(cellIndex);
	}
	
	public static String getValue() {
		return CellReferenceUtil.getValue();
	}
	
}
