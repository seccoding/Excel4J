package io.github.seccoding.excel.util.write;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.util.ZipSecureFile;

import io.github.seccoding.excel.util.write.share.WriteShare;

public class WriteFileSystem {

	public static String write() {
		FileOutputStream fos = null;
		String downloadPath = null;
		
		try {
			
			downloadPath = WriteShare.writeOption.getFilePath();
			if ( downloadPath == null ) {
				throw new RuntimeException("Excel 파일이 만들어질 경로가 누락되었습니다. WriteOption 의 filePath를 셋팅하세요. 예 > D:\\uploadFiles\\");
			}
			
			ZipSecureFile.setMinInflateRatio(0);
			fos = new FileOutputStream(downloadPath + WriteShare.writeOption.getFileName());
			WriteShare.wb.write(fos);
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
