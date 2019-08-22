package io.github.seccoding.excel.util.write;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import io.github.seccoding.excel.annotations.Title;
import io.github.seccoding.excel.util.write.share.WriteShare;

public class MakeContents {

	public static void make() {
		Row row = null;

		List<?> values = WriteShare.writeOption.getContents();

		if (values != null && values.size() > 0) {
			MakeCell makeCell = null;
			while ( true ) {
				Object obj = getValue(values);
				if ( obj == null ) {
					break;
				}
				
				row = MakeRow.create();
				makeCellAndFillValue(obj, makeCell, row);
				values.remove(0);
				flush();
			}
		}
	}
	
	private static Object getValue(List<?> values) {
		try {
			return values.get(0);
		}
		catch ( IndexOutOfBoundsException e ) {
			return null;
		}
	}
	
	private static void makeCellAndFillValue(Object obj, MakeCell makeCell, Row row) {
		
		int cellIndex = 0;
		
		java.lang.reflect.Field[] fields = obj.getClass().getDeclaredFields();
		for (java.lang.reflect.Field field : fields) {
			field.setAccessible(true);

			if (field.isAnnotationPresent(Title.class)) {
				Title anno = field.getAnnotation(Title.class);

				String title = anno.value();
				cellIndex = getColumnIndex(title);
				if ( makeCell == null ) {
					makeCell = new MakeCell(obj, anno, row, cellIndex);
				}
				else {
					makeCell.changeCell(obj, anno, row, cellIndex);
				}
				makeCell.fillValue(field);
			}
		}
	}
	
	private static void flush() {
		if ( WriteShare.wb instanceof SXSSFWorkbook ) {
			try {
				((SXSSFSheet)WriteShare.sheet).flushRows(10000);
			} catch (IOException e) {
				throw new RuntimeException(e.getMessage(), e);
			}
		}
	}
	
	private static int getColumnIndex(String title) {
		return WriteShare.writeOption.getTitles().indexOf(title);
	}
	
}
