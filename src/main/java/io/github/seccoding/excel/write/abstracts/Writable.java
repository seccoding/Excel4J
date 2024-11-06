package io.github.seccoding.excel.write.abstracts;

import io.github.seccoding.excel.annotations.ExcelSheet;

/**
 * 엑셀 파일을 씀.
 * @param <T> 데이터가 들어있는 인스턴스의 원본 클래스
 */
public abstract class Writable<T> {

	/**
	 * 데이터가 들어있는 인스턴스의 원본 클래스
	 */
	protected Class<T> dataClass;
	/**
	 * 워크시트에 작성할 시트 명
	 */
	protected String sheetName;
	/**
	 * 데이터를 쓸 시작 행번호
	 */
	protected int writeStartRow;
	
	protected Writable(Class<T> dataClass) {
		this.dataClass = dataClass;
		this.extractSheetName(dataClass);
	}
	
	/**
	 * 시트 명을 @ExcelSheet에서 추출
	 */
	protected String extractSheetName(Class<?> dataClass) {
		if (dataClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet es = dataClass.getAnnotation(ExcelSheet.class);
			this.sheetName = es.value() == null || es.value().length() == 0 ? "Sheet1" : es.value();
			this.writeStartRow = es.startRow() - 1;
		}
		else {
			this.sheetName = "Sheet1";
			this.writeStartRow = 0;
		}
		
		if (this.writeStartRow < 0) this.writeStartRow = 0;
		
		return this.sheetName;
	}
	
}
