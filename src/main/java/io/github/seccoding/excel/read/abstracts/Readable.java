package io.github.seccoding.excel.read.abstracts;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class Readable<T> {

	protected Workbook workbook;
	protected Sheet sheet;
	protected Row row;
	protected Cell cell;
	
	protected Class<T> resultClass;
	
	protected Readable(Class<T> resultClass) {
		this.resultClass = resultClass;
	}
	
	protected Field getField(String fieldName) {
		try {
			Field field = resultClass.getDeclaredField(fieldName);
			field.setAccessible(true);
			return field;
		} catch (NoSuchFieldException | SecurityException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	protected boolean isPresentFieldAnnotation(String fieldName) {
		Field field = this.getField(fieldName);
		return isAnnotationPresent(field, io.github.seccoding.excel.annotations.Field.class);
	}
	
	protected boolean isPresentFieldAnnotation(Field field) {
		return isAnnotationPresent(field, io.github.seccoding.excel.annotations.Field.class);
	}
	
	protected io.github.seccoding.excel.annotations.Field getFieldAnnotation(Field field) {
		if (isAnnotationPresent(field, io.github.seccoding.excel.annotations.Field.class)) {
			return field.getDeclaredAnnotation(io.github.seccoding.excel.annotations.Field.class);
		}
		return null;
	}
	
	protected boolean isAnnotationPresent(Field field, Class<? extends Annotation> annotationClass) {
		return field.isAnnotationPresent(annotationClass);
	}
}
