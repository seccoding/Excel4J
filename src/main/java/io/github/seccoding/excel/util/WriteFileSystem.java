package io.github.seccoding.excel.util;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.option.WriteOption;

public class WriteFileSystem {

	public static String write(WriteOption<?> writeOption, Workbook wb) {
		FileOutputStream fos = null;
		String downloadPath = null;
		
		try {
			
			downloadPath = writeOption.getFilePath();
			if ( downloadPath == null ) {
				throw new RuntimeException("Excel 파일이 만들어질 경로가 누락되었습니다. WriteOption 의 filePath를 셋팅하세요. 예 > D:\\uploadFiles\\");
			}
			
			fos = new FileOutputStream(downloadPath + writeOption.getFileName());
			wb.write(fos);
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
		}
		
		return downloadPath;
	}
	
}
