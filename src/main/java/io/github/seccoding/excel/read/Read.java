package io.github.seccoding.excel.read;

import java.nio.file.Path;
import java.util.List;

import io.github.seccoding.excel.read.abstracts.ReadWorkbook;

/**
 * 엑셀 워크시트의 내용을 읽음.
 * @param <T>
 */
public class Read<T> extends ReadWorkbook<T> {

	/**
	 * @param excelFilePath 읽으려는 엑셀파일의 경로
	 * @param resultClass 엑셀 내용을 담으려는 클래스.
	 */
	public Read(Path excelFilePath, Class<T> resultClass) {
		super(resultClass);
		super.loadWorkbook(excelFilePath);
	}
	
	/**
	 * 엑셀 파일의 내용을 읽어냄.
	 * @return
	 */
	public List<T> read() {
		return super.setValueInExcel();
	}
	
}
