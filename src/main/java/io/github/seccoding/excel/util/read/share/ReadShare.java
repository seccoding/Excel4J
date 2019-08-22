package io.github.seccoding.excel.util.read.share;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.option.ReadOption;

public class ReadShare {

	public static Workbook wb;
	
	public static Sheet sheet;
	public static String sheetName;
	
	public static Row row;
	public static Cell cell;
	
	public static Class<?> clazz;
	
	public static int numOfRows;
	public static int numOfCells;
	
	public static ReadOption readOption;
}
