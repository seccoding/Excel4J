package io.github.seccoding.excel.write.abstracts;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import io.github.seccoding.excel.annotations.Align;
import io.github.seccoding.excel.annotations.BackgroundColor;
import io.github.seccoding.excel.annotations.Border;
import io.github.seccoding.excel.annotations.Text;
import io.github.seccoding.excel.annotations.Title;
import io.github.seccoding.excel.util.FileType;

public abstract class WriteWorkbook<T> extends Writable<T> {

	protected Workbook workbook;
	protected Sheet sheet;
	protected CellStyle borderStyle;
	protected Map<String, CellStyle> backgroundStyle;
	protected Map<String, CellStyle> textStyle;
	protected Map<String, CellStyle> alignStyle;

	protected WriteWorkbook(Class<T> dataClass) {
		super(dataClass);
	}

	protected void makeWorkbook(String fileName) {
		if (FileType.isXls(fileName)) {
			this.workbook = new HSSFWorkbook();
		} else if (FileType.isXlsx(fileName)) {
			this.workbook = new SXSSFWorkbook(-1);
		} else {
			throw new RuntimeException("Could not find Excel file");
		}
	}

	protected void makeSheet() {
		this.sheet = this.workbook.createSheet(super.sheetName);
		this.makeBorder();
		this.makeBackgroundColor();
		this.makeTextColor();
		this.makeAlign();
	}

	protected void autoColumnSize() {
		if (this.sheet instanceof SXSSFSheet) {
			((SXSSFSheet) this.sheet).trackAllColumnsForAutoSizing();
		}

		Row row = this.sheet.getRow(this.sheet.getFirstRowNum());

		for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			this.sheet.autoSizeColumn(j);
		}
	}

	private void makeBorder() {
		if (super.dataClass.isAnnotationPresent(Border.class)) {
			Border border = super.dataClass.getAnnotation(Border.class);
			this.borderStyle = this.workbook.createCellStyle();

			if (border.top()) {
				borderStyle.setBorderTop(border.value());
				borderStyle.setTopBorderColor(border.color().index);
			}
			if (border.right()) {
				borderStyle.setBorderRight(border.value());
				borderStyle.setRightBorderColor(border.color().index);
			}
			if (border.bottom()) {
				borderStyle.setBorderBottom(border.value());
				borderStyle.setBottomBorderColor(border.color().index);
			}
			if (border.left()) {
				borderStyle.setBorderLeft(border.value());
				borderStyle.setLeftBorderColor(border.color().index);
			}
		}
	}

	protected void makeBackgroundColor() {
		this.backgroundStyle = new HashMap<>();
		
		Field[] fields = super.dataClass.getDeclaredFields();
		for (Field field : fields) {

			if (field.isAnnotationPresent(Title.class) 
					&& field.isAnnotationPresent(BackgroundColor.class)) {
				
				BackgroundColor bg = field.getAnnotation(BackgroundColor.class);
				
				CellStyle backgroundColor = this.workbook.createCellStyle();
				backgroundColor.cloneStyleFrom(this.borderStyle);
				backgroundColor.setFillForegroundColor(bg.value().index);
				backgroundColor.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				
				this.backgroundStyle.put(field.getName(), backgroundColor);
			}
		}
	}
	
	protected void makeTextColor() {
		this.textStyle = new HashMap<>();
		
		Field[] fields = super.dataClass.getDeclaredFields();
		for (Field field : fields) {
			
			if (field.isAnnotationPresent(Title.class) 
					&& field.isAnnotationPresent(Text.class)) {
				Text text = field.getAnnotation(Text.class);
				
				CellStyle textColor = this.workbook.createCellStyle();
				
				if (this.backgroundStyle.containsKey(field.getName())) {
					textColor.cloneStyleFrom(this.backgroundStyle.get(field.getName()));
				}
				else {
					textColor.cloneStyleFrom(this.borderStyle);
				}
				
				Font font = this.workbook.createFont();
				font.setBold(text.bold());
				font.setColor(text.color().index);
				textColor.setFont(font);
				
				this.textStyle.put(field.getName(), textColor);
			}
		}
	}
	
	protected void makeAlign() {
		this.alignStyle = new HashMap<>();
		
		Field[] fields = super.dataClass.getDeclaredFields();
		for (Field field : fields) {
			
			if (field.isAnnotationPresent(Title.class) 
					&& field.isAnnotationPresent(Align.class)) {
				Align align = field.getAnnotation(Align.class);
				
				CellStyle textAlign = this.workbook.createCellStyle();
				
				if (this.textStyle.containsKey(field.getName())) {
					textAlign.cloneStyleFrom(this.textStyle.get(field.getName()));
				}
				else if (this.backgroundStyle.containsKey(field.getName())) {
					textAlign.cloneStyleFrom(this.backgroundStyle.get(field.getName()));
				}
				else {
					textAlign.cloneStyleFrom(this.borderStyle);
				}
				
				textAlign.setAlignment(align.value());
				textAlign.setVerticalAlignment(align.verticalAlignment());
				
				this.alignStyle.put(field.getName(), textAlign);
			}
		}
	}
}
