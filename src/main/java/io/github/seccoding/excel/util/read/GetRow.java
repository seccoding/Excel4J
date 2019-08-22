package io.github.seccoding.excel.util.read;

import io.github.seccoding.excel.util.read.share.ReadShare;

public class GetRow {

	public static void setRow(int rowIndex) {
		ReadShare.row = ReadShare.sheet.getRow(rowIndex);
	}
	
	public static boolean isNotNull() {
		return ReadShare.row != null;
	}
	
	public static void setPhysicalNumberOfCells() {
		ReadShare.numOfCells = ReadShare.row.getPhysicalNumberOfCells();
	}
	
}
