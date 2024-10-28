package io.github.seccoding.excel.write.abstracts;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import io.github.seccoding.excel.annotations.Title;
import io.github.seccoding.excel.util.InstanceUtil;

public abstract class WriteBody<T> extends WriteTitle<T> {

	protected WriteBody(Class<T> dataClass, List<T> contents) {
		super(dataClass, contents);
	}

	protected void makeContentRow() {
		for (T t : contents) {
			Row row = super.sheet.createRow(this.nextRowIndex);
			int cellIndex = 0;
			
			Field[] fields = t.getClass().getDeclaredFields();
			for (Field field : fields) {
				if (field.isAnnotationPresent(Title.class)) {
					
					Cell cell = row.createCell(cellIndex);
					super.setCellStyle(cell, field);
					
					Method getter = InstanceUtil.getMethod(t, "get", field.getName());
					Object value = InstanceUtil.invokeMethod(t, getter);
					
					Class<?> fieldType = field.getType();
					if (fieldType == String.class) {
						String cellValue = value.toString();
						if (cellValue.startsWith("=")) {
							cell.setCellFormula(cellValue);
						}
						else {
							cell.setCellValue(cellValue);
						}
					}
					else if (fieldType == int.class) {
						cell.setCellValue(Integer.parseInt(value.toString()));
					}
					else if (fieldType == double.class) {
						cell.setCellValue(Double.parseDouble(value.toString()));
					}
					else if (fieldType == boolean.class) {
						cell.setCellValue(Boolean.parseBoolean(value.toString()));
					}
					
					cellIndex++;
				}
			}
			
			
			this.nextRowIndex++;
		}
	}
}
