package io.github.seccoding.excel.util.write;

import java.lang.reflect.Field;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import io.github.seccoding.excel.annotations.Title;

public class MakeParentTitle {

	public static int make(Title title, Row row, Cell cell, int cellIndex, Field[] fields) {
		if ( !title.parentTitle().equals("") ) {
			if ( isMergeCellOrRow(title) ) {
				String parentTitle = replaceTitle(title.parentTitle(), fields);
				cell = row.createCell(cellIndex);
				cell.setCellValue(parentTitle);
				
				CellMerger.merge(row.getRowNum(), title.parentRowMerge(), cellIndex, title.parentCellMerge());
				
				Row tempRow = MakeRow.create(row.getRowNum() + 1, true);
				
				cell = tempRow.createCell(cellIndex);
				cell.setCellValue(title.value());
				
				CellMerger.merge(tempRow.getRowNum(), title.rowMerge(), cellIndex, title.cellMerge());
				cellIndex += title.cellMerge();
			}
			
			if ( isNotMergeCellOrRow(title) ) {
				
				String titleValue = replaceTitle(title.value(), fields);
				
				Row tempRow = MakeRow.create(row.getRowNum() + 1);
				cell = tempRow.createCell(cellIndex);
				cell.setCellValue(titleValue);
				
				CellMerger.merge(tempRow.getRowNum(), title.rowMerge(), cellIndex, title.cellMerge());
				cellIndex += title.cellMerge();
			}
			
		}
		
		return cellIndex;
	}
	
	private static String replaceTitle(String title, Field[] fields) {
		if ( title.startsWith("$") ) {
			String fieldName = title.replace("$", "");
			title = GetHeaderContent.getParentTitle(GetHeaderContent.getFirstContent(), fields, fieldName);
		}
		
		return title;
	}
	
	private static boolean isMergeCellOrRow(Title title) {
		return title.parentCellMerge() > 1 || title.parentRowMerge() > 1;
	}
	
	private static boolean isNotMergeCellOrRow(Title title) {
		return title.parentCellMerge() == 1 && title.parentRowMerge() == 1;
	}
	
}
