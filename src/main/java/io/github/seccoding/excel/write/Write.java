package io.github.seccoding.excel.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.util.ZipSecureFile;

import io.github.seccoding.excel.write.abstracts.WriteBody;

public class Write<T> extends WriteBody<T> {

	public Write(Class<T> dataClass, List<T> contents) {
		super(dataClass, contents);
	}
	
	public void write(File excelPath) {
		super.makeWorkbook(excelPath.getName());
		super.makeSheet();
		
		super.makeMainTitleRow();
		super.makeContentRow();
		super.autoColumnSize();
		this.writeFile(excelPath);
	}
	
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
