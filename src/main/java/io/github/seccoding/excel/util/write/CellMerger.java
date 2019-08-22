package io.github.seccoding.excel.util.write;

import io.github.seccoding.excel.util.write.share.WriteShare;

public class CellMerger {

	public static void merge(int fromRow, int toRow, int fromCell, int toCell) {
		if ( isExtends(fromRow, toRow, fromCell, toCell) ) {
			WriteShare.sheet.addMergedRegion(CellRangeChanger.cellRangeAddress(fromRow, toRow, fromCell, toCell));
		}
	}
	
	private static boolean isExtends(int fromRow, int toRow, int fromCell, int toCell) {
		return fromRow < fromRow + toRow - 1 || fromCell < fromCell + toCell - 1;
	}
	
}
