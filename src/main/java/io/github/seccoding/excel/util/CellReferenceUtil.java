package io.github.seccoding.excel.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

public class CellReferenceUtil {

	public static String getName(Cell cell, int cellIndex) {
		return CellReference.convertNumToColString( cell != null ? cell.getColumnIndex() : cellIndex);
	}
	
	public static String getValue(Cell cell) {
		if(cell == null) {
			return "";
		}
		if( cell.getCellType() == Cell.CELL_TYPE_FORMULA ) {
			return cell.getCellFormula();
		}
		if( cell.getCellType() == Cell.CELL_TYPE_NUMERIC ) {
			return cell.getNumericCellValue() + "";
		}
		if( cell.getCellType() == Cell.CELL_TYPE_STRING ) {
			return cell.getStringCellValue();
		}
		if( cell.getCellType() == Cell.CELL_TYPE_BOOLEAN ) {
			return cell.getBooleanCellValue() + "";
		}
		if( cell.getCellType() == Cell.CELL_TYPE_ERROR ) {
			return cell.getErrorCellValue() + "";
		}
		if( cell.getCellType() == Cell.CELL_TYPE_BLANK ) {
			return "";
		}
		
		return cell.getStringCellValue();
	}
	
}
