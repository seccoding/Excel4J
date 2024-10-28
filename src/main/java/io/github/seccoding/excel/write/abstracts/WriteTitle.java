package io.github.seccoding.excel.write.abstracts;

import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import io.github.seccoding.excel.annotations.Title;

public abstract class WriteTitle<T> extends WriteWorkbook<T> {

	protected List<T> contents;
	protected int nextRowIndex;
	
	protected WriteTitle(Class<T> dataClass, List<T> contents) {
		super(dataClass);
		this.contents = contents;
	}

	protected void makeMainTitleRow() {

		Row row = super.sheet.createRow(super.writeStartRow);
		int cellIndex = 0;
		
		Field[] fields = super.dataClass.getDeclaredFields();
		for (Field field : fields) {
			if (field.isAnnotationPresent(Title.class)) {
				Title title = field.getAnnotation(Title.class);
				
				Cell cell = row.createCell(cellIndex);
				cell.setCellValue(title.value());
				
				if (super.borderStyle != null) {
					cell.setCellStyle(super.borderStyle);
				}
				if (super.backgroundStyle.containsKey(field.getName())) {
					cell.setCellStyle(super.backgroundStyle.get(field.getName()));
				}
				if (super.textStyle.containsKey(field.getName())) {
					cell.setCellStyle(super.textStyle.get(field.getName()));
				}
				if (super.alignStyle.containsKey(field.getName())) {
					cell.setCellStyle(super.alignStyle.get(field.getName()));
				}
				cellIndex++;
			}
		}
		
		this.nextRowIndex = this.sheet.getPhysicalNumberOfRows();
	}
	
}
