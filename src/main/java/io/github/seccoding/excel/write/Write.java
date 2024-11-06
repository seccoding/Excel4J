package io.github.seccoding.excel.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Sheet;

import io.github.seccoding.excel.write.abstracts.WriteBody;

/**
 * 엑셀 파일을 작성한다.
 * @param <T> 엑셀 파일에 작성할 데이터가 들어있는 인스턴스의 원본 클래스
 */
public class Write<T> extends WriteBody<T> {

	/**
	 * @param dataClass 엑셀 파일에 작성할 데이터가 들어있는 인스턴스의 원본 클래스
	 * @param contents 엑셀 파일에 작성할 리스트 인스턴스
	 */
	public Write(Class<T> dataClass, List<T> contents) {
		super(dataClass, contents);
	}
	
	/**
	 * 엑셀파일에 새로운 시트를 생성하고 데이터를 작성한다.
	 * @param sheetName 새로운 시트명
	 * @param data 새로운 시트에 작성할 리스트 인스턴스
	 */
	public void appendNewSheet(String sheetName, List<T> data) {
		super.extractSheetName(super.dataClass); 
		this.appendNewSheet(sheetName, super.writeStartRow, super.dataClass, data);
	}
	
	/**
	 * 엑셀파일에 새로운 시트를 생성하고 데이터를 작성한다.
	 * @param dataClass 새로운 시트에 작성할 데이터가 들어있는 인스턴스의 원본 클래스
	 * @param data 새로운 시트에 작성할 리스트 인스턴스
	 */
	public void appendNewSheet(Class<?> dataClass, List<? extends Object> data) {
		String sheetName = super.extractSheetName(dataClass);
		this.appendNewSheet(sheetName, super.writeStartRow, dataClass, data);
	}
	
	/**
	 * 엑셀파일에 새로운 시트를 생성하고 데이터를 작성한다.
	 * @param sheetName 새로운 시트명
	 * @param dataClass 새로운 시트에 작성할 데이터가 들어있는 인스턴스의 원본 클래스
	 * @param data 새로운 시트에 작성할 리스트 인스턴스
	 */
	public void appendNewSheet(String sheetName, Class<?> dataClass, List<? extends Object> data) {
		super.extractSheetName(dataClass);
		this.appendNewSheet(sheetName, super.writeStartRow, dataClass, data);
	}
	
	/**
	 * 엑셀파일에 새로운 시트를 생성하고 데이터를 작성한다.
	 * @param writeStartRow 새로운 시트에 작성할 행의 시작 번호
	 * @param dataClass 새로운 시트에 작성할 데이터가 들어있는 인스턴스의 원본 클래스
	 * @param data 새로운 시트에 작성할 리스트 인스턴스
	 */
	public void appendNewSheet(int writeStartRow, Class<?> dataClass, List<? extends Object> data) {
		String sheetName = super.extractSheetName(dataClass);
		super.extractSheetName(dataClass);
		this.appendNewSheet(sheetName, writeStartRow, dataClass, data);
	}
	
	/**
	 * 엑셀파일에 새로운 시트를 생성하고 데이터를 작성한다.
	 * @param sheetName 새로운 시트명
	 * @param writeStartRow 새로운 시트에 작성할 행의 시작 번호
	 * @param dataClass 새로운 시트에 작성할 데이터가 들어있는 인스턴스의 원본 클래스
	 * @param data 새로운 시트에 작성할 리스트 인스턴스
	 */
	public void appendNewSheet(String sheetName, int writeStartRow, Class<?> dataClass, List<? extends Object> data) {
		Sheet sheet = super.makeSheet(sheetName, dataClass);
		
		super.makeMainTitleRow(sheet, dataClass, writeStartRow);
		super.makeContentRow(sheet, data);
		super.autoColumnSize(sheet);
	}
	
	/**
	 * 엑셀파일을 생성하고 내용을 작성한다.
	 * @param filename 엑셀 파일의 이름
	 */
	public void write(String filename) {
		super.makeWorkbook(filename);
		Sheet sheet = super.makeSheet();
		
		super.makeMainTitleRow(sheet, super.dataClass, super.writeStartRow);
		super.makeContentRow(sheet, super.contents);
		super.autoColumnSize(sheet);
	}
	
	public void toFile(File excelPath) {
		this.writeFile(excelPath);
	}
	
	/**
	 * 엑셀 워트북을 파일 인스턴스로 변환 (디스크에 엑셀 파일을 작성한다)
	 * @param excelPath
	 */
	private void writeFile(File excelPath) {
		FileOutputStream fos = null;
		
		try {
			if ( excelPath == null ) {
				throw new RuntimeException("Excel 파일이 만들어질 경로가 누락되었습니다.");
			}
			
			ZipSecureFile.setMinInflateRatio(0);
			fos = new FileOutputStream(excelPath);
			super.workbook.write(fos);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
		finally {
			if(fos != null) {
				try {
					fos.flush();
				} catch (IOException e) {}
				
				try {
					fos.close();
				} catch (IOException e) {}
			}
			
			try {
				super.workbook.close();
			} catch (IOException e) {
			}
		}
	}

}
