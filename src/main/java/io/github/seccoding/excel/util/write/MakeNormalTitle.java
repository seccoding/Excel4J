package io.github.seccoding.excel.util.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import io.github.seccoding.excel.annotations.Title;

public class MakeNormalTitle {

	public static int make(Title title, Row row, Cell cell, int cellIndex) {
		if ( title.parentTitle().equals("") ) {
			cell = row.createCell(cellIndex);
			
			cell.setCellValue(title.value());
			CellMerger.merge(row.getRowNum(), title.rowMerge(), cellIndex, title.cellMerge());
			cellIndex += title.cellMerge();
		}
		
		return cellIndex;
	}
	
}
