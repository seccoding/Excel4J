package io.github.seccoding.excel.util.write.share;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.option.WriteOption;

public class WriteShare {

	public static WriteOption<?> writeOption;
	
	
	public static Workbook wb;
	
	/**
	 * 엑셀 문서에 만들어질 Sheet
	 */
	public static Sheet sheet;

	/**
	 * 엑셀 문서에 Row를 작성할 때 몇 번째에 Row를 만들 것인지 지정하기 위한 변수 엑셀 문서에 Row를 작성할 때마다 증가함.
	 */
	public static int rowIndex;
	
	
	public static void resetRowIndex() {
		rowIndex = 0;
	}
	
}
