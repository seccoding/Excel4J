package io.github.seccoding.excel.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.util.ZipSecureFile;

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
	 * 엑셀파일을 생성하고 내용을 작성한다.
	 * @param excelPath 생성할 엑셀 파일의 파일 인스턴스
	 */
	public void write(File excelPath) {
		super.makeWorkbook(excelPath.getName());
		super.makeSheet();
		
		super.makeMainTitleRow();
		super.makeContentRow();
		super.autoColumnSize();
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
				throw new RuntimeException("Excel 파일이 만들어질 경로가 누락되었습니다. WriteOption 의 filePath를 셋팅하세요. 예 > D:\\uploadFiles\\");
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
