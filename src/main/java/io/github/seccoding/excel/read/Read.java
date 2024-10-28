package io.github.seccoding.excel.read;

import java.nio.file.Path;
import java.util.List;

import io.github.seccoding.excel.read.abstracts.ReadWorkbook;

public class Read<T> extends ReadWorkbook<T> {

	public Read(Path excelFilePath, Class<T> resultClass) {
		super(resultClass);
		super.loadWorkbook(excelFilePath);
	}
	
	public List<T> read() {
		return super.setValueInExcel();
	}
	
}
