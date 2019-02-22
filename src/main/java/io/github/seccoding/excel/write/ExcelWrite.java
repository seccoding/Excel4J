package io.github.seccoding.excel.write;

import java.io.File;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.option.WriteOption;
import io.github.seccoding.excel.util.MakeWorkBook;
import io.github.seccoding.excel.util.WriteFileSystem;

/**
 * 엑셀 파일을 서버의 디스크에 작성한다.
 * 
 * @see io.github.seccoding.excel.ExcelWriteTest
 * @author Minchang Jang (mcjang1116@gmail.com)
 */
public class ExcelWrite {

	/**
	 * 엑셀 파일이 쓰여질 경로. WriteOption 에서 가져온다.
	 */
	private static String downloadPath = null;

	private static Workbook wb;
	
	/**
	 * 엑셀 문서에 만들어질 Sheet
	 */
	private static Sheet sheet;

	/**
	 * 엑셀 문서에 Row를 작성할 때 몇 번째에 Row를 만들 것인지 지정하기 위한 변수 엑셀 문서에 Row를 작성할 때마다 증가함.
	 */
	private static int rowIndex;

	/**
	 * 엑셀 파일을 작성한다.
	 * 
	 * @param WriteOption
	 * @return Excel 파일의 File 객체
	 */
	public static File write(WriteOption<?> writeOption) {

		wb = MakeWorkBook.getWorkbook(writeOption.getFileName());
		sheet = wb.createSheet(writeOption.getSheetName());
		setTitle(writeOption.getTitles());
		setContents(writeOption);

		downloadPath = WriteFileSystem.write(writeOption, wb);
		rowIndex = 0;

		return getFile(writeOption.getFileName());
	}

	private static void setTitle(List<String> values) {

		Row row = null;
		Cell cell = null;

		int cellIndex = 0;

		if (values != null && values.size() > 0) {
			row = sheet.createRow(rowIndex++);
			for (String value : values) {
				cell = row.createCell(cellIndex++);
				cell.setCellValue(value);
			}
		}

	}

	private static void setContents(WriteOption<?> writeOption) {

		Row row = null;

		List<?> values = writeOption.getContents();

		int cellIndex = 0;
		if (values != null && values.size() > 0) {
			
			while ( true ) {
				Object obj = null;
				try {
					obj = values.get(0);
				}
				catch ( IndexOutOfBoundsException e ) {
					break;
				}
				
				if ( obj == null ) {
					break;
				}
				
				row = sheet.createRow(rowIndex++);
				cellIndex = 0;

				java.lang.reflect.Field[] fields = obj.getClass().getDeclaredFields();
				for (java.lang.reflect.Field f : fields) {
					f.setAccessible(true);

					if (f.isAnnotationPresent(Field.class)) {
						Field anno = f.getAnnotation(Field.class);

						String title = anno.value();
						cellIndex = getColumnIndex(title, writeOption);
						MakeCell makeCell = new MakeCell(obj, anno, row, cellIndex);
						makeCell.fillValue(wb, sheet, f);
					}

				}
				
				values.remove(0);
			}
			
		}

	}

	private static int getColumnIndex(String title, WriteOption<?> writeOption) {
		return writeOption.getTitles().indexOf(title);
	}

	
	private static File getFile(String fileName) {
		return new File(downloadPath + fileName);
	}

}
