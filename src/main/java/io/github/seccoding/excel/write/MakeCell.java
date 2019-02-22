package io.github.seccoding.excel.write;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Format;

public class MakeCell {

	private static List<Class<?>> numericTypes;

	static {
		numericTypes = new ArrayList<Class<?>>();
		numericTypes.add(Byte.class);
		numericTypes.add(Short.class);
		numericTypes.add(Integer.class);
		numericTypes.add(Long.class);
		numericTypes.add(Float.class);
		numericTypes.add(Double.class);
	}
	
	private Object obj;
	private Field fieldAnnotation;
	private Format format;
	private Row row;
	private int cellIndex;

	public MakeCell(Object obj, Field fieldAnnotation, Row row, int cellIndex) {
		this.obj = obj;
		this.fieldAnnotation = fieldAnnotation;
		this.row = row;
		this.cellIndex = cellIndex;
	}
	
	public void fillValue(Workbook wb, Sheet sheet, java.lang.reflect.Field f) {
		Cell cell = null;
		try {
			obj = f.get(obj);
			
			format = f.getAnnotation(Format.class);
			
			cell = makeCellAndFill();
			CellStyle cellStyle = makeCellStyle(wb);
			
			if ( cell != null ) {
				sheet.autoSizeColumn(cellIndex);
				cell.setCellStyle(cellStyle);
			}
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private CellStyle makeCellStyle(Workbook wb) {
		
		CellStyle style = wb.createCellStyle();
		style.setAlignment(format.alignment());
		style.setVerticalAlignment(format.verticalAlignment());
		
		String formatString = format.dataFormat();
		if ( !fieldAnnotation.date() && formatString != null && formatString.length() > 0 ) {
			DataFormat dataFormat = wb.createDataFormat();
			style.setDataFormat(dataFormat.getFormat(formatString));
		}
		
		if ( format.bold() ) {
			Font font = wb.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			style.setFont(font);
		}
		
		return style;
	}
	
	private Cell makeCellAndFill() {
		Cell cell = null;
		if (obj.getClass() == String.class) {
			
			String data = obj + "";
			if ( fieldAnnotation.date() ) {
				data = makeDateTime(data);
				cell = row.createCell(cellIndex);
				cell.setCellValue(data);
			}
			else if (data.trim().startsWith("=")) {
				data = data.trim().substring(1).trim();
				cell = row.createCell(cellIndex, Cell.CELL_TYPE_FORMULA);
				cell.setCellFormula(data);
			} else {
				cell = row.createCell(cellIndex, Cell.CELL_TYPE_STRING);
				cell.setCellValue(data);
			}
			
		} else if (numericTypes.contains(obj.getClass())) {
			cell = row.createCell(cellIndex, Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue(Double.parseDouble(String.valueOf(obj)));
		} else if (obj.getClass() == Boolean.class) {
			cell = row.createCell(cellIndex, Cell.CELL_TYPE_BOOLEAN);
			cell.setCellValue(Boolean.parseBoolean(obj + ""));
		}
		
		return cell;
	}
	
	private String makeDateTime(String data) {
		String formatString = format.dataFormat();
		if ( formatString == null || formatString.length() == 0 ) {
			throw new RuntimeException("dataFormat이 지정되지 않았습니다.");
		}
		
		formatString = formatString.trim();
		
		try {
			Date date = new SimpleDateFormat(formatString).parse(data.trim());
			String toDataFormat = format.toDataFormat();
			if ( toDataFormat != null && toDataFormat.length() > 0  ) {
				formatString = toDataFormat.trim();
			}
			return new SimpleDateFormat(formatString).format(date);
		} catch (ParseException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
}
