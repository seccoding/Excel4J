package io.github.seccoding.excel.util.write;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import io.github.seccoding.excel.util.write.share.WriteShare;

public class AutoSizingColumns {

	public static void resize() {
		if (WriteShare.sheet instanceof SXSSFSheet) {
			((SXSSFSheet) WriteShare.sheet).trackAllColumnsForAutoSizing();
		}

		Row row = WriteShare.sheet.getRow(0);
		
		for ( int j = 0; j < row.getLastCellNum(); j++ ) {
			WriteShare.sheet.autoSizeColumn(j);
		}
	}
	
}
