package io.github.seccoding.excel.util.write;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.annotations.Format;
import io.github.seccoding.excel.annotations.Title;
import io.github.seccoding.excel.util.write.share.WriteShare;

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
	private Title fieldAnnotation;
	private Format format;
	private Row row;
	private int cellIndex;
	private CellStyle style;
	private DataFormat dataFormat;
	private Font font;
	
	public MakeCell(Object obj, Title fieldAnnotation, Row row, int cellIndex) {
		this.obj = obj;
		this.fieldAnnotation = fieldAnnotation;
		this.row = row;
		this.cellIndex = cellIndex;
	}
	
	public void changeCell(Object obj, Title fieldAnnotation, Row row, int cellIndex) {
		this.obj = obj;
		this.fieldAnnotation = fieldAnnotation;
		this.row = row;
		this.cellIndex = cellIndex;
	}
	
	public void fillValue(java.lang.reflect.Field f) {
		Cell cell = null;
		try {
			obj = f.get(obj);
			
			format = f.getAnnotation(Format.class);
			
			cell = makeCellAndFill();
			CellStyle cellStyle = makeCellStyle(WriteShare.wb);
			
			if ( cell != null ) {
				cell.setCellStyle(cellStyle);
			}
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private Cell makeCellAndFill() {
		Cell cell = null;
		
		if ( obj == null ) {
			cell = row.createCell(cellIndex, CellType.STRING);
			cell.setCellValue("");
			return cell;
		}
		
		if (obj.getClass() == String.class) {
			
			String data = obj + "";
			if ( fieldAnnotation.date() ) {
				data = makeDateTime(data);
				cell = row.createCell(cellIndex);
				cell.setCellValue(data);
			}
			else if (data.trim().startsWith("=")) {
				data = data.trim().substring(1).trim();
				cell = row.createCell(cellIndex, CellType.FORMULA);
				cell.setCellFormula(data);
			} else {
				cell = row.createCell(cellIndex, CellType.STRING);
				cell.setCellValue(data);
			}
			
		} else if (numericTypes.contains(obj.getClass())) {
			cell = row.createCell(cellIndex, CellType.NUMERIC);
			cell.setCellValue(Double.parseDouble(String.valueOf(obj)));
		} else if (obj.getClass() == Boolean.class) {
			cell = row.createCell(cellIndex, CellType.BOOLEAN);
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
		
		if ( data == null || data.trim().length() == 0 || data.contains("null")) {
			return "";
		}
		
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
	
	private CellStyle makeCellStyle(Workbook wb) {
		
		//if ( style == null ) {
			style = wb.createCellStyle();
		//}
		
		String alignment = format.alignment();
		if ( alignment.equals(Format.LEFT) ) {
			style.setAlignment(HorizontalAlignment.LEFT);
		}
		else if ( alignment.equals(Format.CENTER) ) {
			style.setAlignment(HorizontalAlignment.CENTER);
		}
		else if ( alignment.equals(Format.RIGHT) ) {
			style.setAlignment(HorizontalAlignment.RIGHT);
		} 
		
		String vAlignment = format.verticalAlignment();
		if ( vAlignment.equals(Format.V_TOP) ) {
			style.setVerticalAlignment(VerticalAlignment.TOP);
		}
		else if ( vAlignment.equals(Format.V_CENTER) ) {
			style.setVerticalAlignment(VerticalAlignment.CENTER);
		}
		else if ( vAlignment.equals(Format.V_BOTTOM) ) {
			style.setVerticalAlignment(VerticalAlignment.BOTTOM);
		} 
		
		String formatString = format.dataFormat();
		if ( !fieldAnnotation.date() && formatString != null && formatString.length() > 0 ) {
			if ( dataFormat == null ) {
				dataFormat = wb.createDataFormat();
			}
			style.setDataFormat(dataFormat.getFormat(formatString));
		}
		
		if ( format.bold() ) {
			if ( font == null ) {
				font = wb.createFont();
			}
			style.setFont(font);
			font.setBold(true);
		}
		
		return style;
	}
	
}
