package io.github.seccoding.excel.util.read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.github.seccoding.excel.util.read.share.ReadShare;
import io.github.seccoding.excel.util.write.FileType;

public class GetWorkbook {

	public static void get(String filePath) {
		ReadShare.wb = getWorkbook(filePath);
	}
	
	public static Workbook getWorkbook(String filePath) {

		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
		
		if (FileType.isXls(filePath)) {
			try {
				return new HSSFWorkbook(fis);
			} catch (IOException e) {
				throw new RuntimeException(e.getMessage(), e);
			} finally {
				if ( fis != null ) {
					try {
						fis.close();
					}
					catch(IOException e1) {}
				}
			}
		}  
		if (FileType.isXlsx(filePath)) {
			try {
				return new XSSFWorkbook(fis);
			} catch (IOException e) {
				throw new RuntimeException(e.getMessage(), e);
			} finally {
				if ( fis != null ) {
					try {
						fis.close();
					}
					catch(IOException e1) {}
				}
			}
		}
		
		if ( fis != null ) {
			try {
				fis.close();
			}
			catch(IOException e1) {}
		}
		
		throw new RuntimeException(filePath + " isn't excel file format");
		
	}
}
