package io.github.seccoding.excel.write.abstracts;

import io.github.seccoding.excel.annotations.ExcelSheet;

public abstract class Writable<T> {

	protected Class<T> dataClass;
	protected String sheetName;
	protected int writeStartRow;
	
	protected Writable(Class<T> dataClass) {
		this.dataClass = dataClass;
		this.extractSheetName();
	}
	
	protected void extractSheetName() {
		if (this.dataClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet es = this.dataClass.getAnnotation(ExcelSheet.class);
			this.sheetName = es.value() == null || es.value().length() == 0 ? "Sheet1" : es.value();
			this.writeStartRow = es.startRow() - 1;
		}
		else {
			this.sheetName = "Sheet1";
			this.writeStartRow = 0;
		}
		
		if (this.writeStartRow < 0) this.writeStartRow = 0;
	}
	
}
