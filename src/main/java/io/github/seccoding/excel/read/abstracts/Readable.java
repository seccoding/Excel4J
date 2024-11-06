package io.github.seccoding.excel.read.abstracts;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 엑셀 워크시트를 읽어냄.
 * @param <T> 워크시트 내용을 할당할 클래스.
 */
public abstract class Readable<T> {

	/**
	 * 읽으려는 엑셀파일의 워크북
	 */
	protected Workbook workbook;
	/**
	 * 읽으려는 엑셀 워크북의 시트.
	 * @ExcelSheet 의 value에 해당하는 시트를 읽는다.
	 */
	protected Sheet sheet;
	
	/**
	 * 읽으려는 엑셀 파일에 존재하는 모든 시트 목록
	 */
	protected List<Sheet> sheetList;
	
	/**
	 * 엑셀 시트 Row의 내용을 할당할 클래스 원본
	 */
	protected Class<T> resultClass;
	
	protected Readable(Class<T> resultClass) {
		this.resultClass = resultClass;
	}
	
	/**
	 * resultClass의 멤버변수를 가져옴. (애노테이션 적용 여부 확인용)
	 * @param fieldName 멤버변수 명
	 * @return 멤버변수 필드
	 */
	protected Field getField(String fieldName) {
		try {
			Field field = resultClass.getDeclaredField(fieldName);
			field.setAccessible(true);
			return field;
		} catch (NoSuchFieldException | SecurityException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	/**
	 * 해당 필드에 @Field가 적용되었는지 확인
	 * @param fieldName 멤버변수 명
	 * @return @Field 적용 여부.
	 */
	protected boolean isPresentFieldAnnotation(String fieldName) {
		Field field = this.getField(fieldName);
		return isAnnotationPresent(field, io.github.seccoding.excel.annotations.Field.class);
	}
	
	/**
	 * 해당 필드에 @Field가 적용되었는지 확인
	 * @param field 멤버변수 필드
	 * @return @Field 적용 여부.
	 */
	protected boolean isPresentFieldAnnotation(Field field) {
		return isAnnotationPresent(field, io.github.seccoding.excel.annotations.Field.class);
	}
	
	/**
	 * resultClass 에서 @Field 가 적용된 멤버변수의 @Field 애노테이션을 조회.
	 * 값 확인 용.
	 * @param field
	 * @return 적용된 @Field 애노테이션
	 */
	protected io.github.seccoding.excel.annotations.Field getFieldAnnotation(Field field) {
		if (isAnnotationPresent(field, io.github.seccoding.excel.annotations.Field.class)) {
			return field.getDeclaredAnnotation(io.github.seccoding.excel.annotations.Field.class);
		}
		return null;
	}
	
	/**
	 * resultClass 의 필드에 해당 애노테이션이 적용되었는지 여부
	 * @param field 멤버변수 필드
	 * @param annotationClass 확인할 애노테이션
	 * @return
	 */
	protected boolean isAnnotationPresent(Field field, Class<? extends Annotation> annotationClass) {
		return field.isAnnotationPresent(annotationClass);
	}
}
