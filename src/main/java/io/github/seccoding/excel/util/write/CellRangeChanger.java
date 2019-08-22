package io.github.seccoding.excel.util.write;

import org.apache.poi.ss.util.CellRangeAddress;

public class CellRangeChanger {

	public static CellRangeAddress cellRangeAddress(int fromRow, int toRow, int fromCell, int toCell) {
		return new CellRangeAddress(fromRow, fromRow + toRow-1, fromCell, fromCell + toCell-1);
	}
	
}
