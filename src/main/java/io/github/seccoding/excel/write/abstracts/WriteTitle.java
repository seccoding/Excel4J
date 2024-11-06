package io.github.seccoding.excel.write.abstracts;

import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import io.github.seccoding.excel.annotations.Merge;
import io.github.seccoding.excel.annotations.Title;

/**
 * 엑셀시트에 타이틀을 작성.
 * @param <T>
 */
public abstract class WriteTitle<T> extends WriteWorkbook<T> {

	/**
	 * 엑셀 시트에 작성할 리스트 인스턴스
	 */
	protected List<T> contents;
	/**
	 * 다음에 작성할 행 번호
	 */
	protected int nextRowIndex;
	
	protected WriteTitle(Class<T> dataClass, List<T> contents) {
		super(dataClass);
		this.contents = contents;
	}

	/**
	 * 타이틀 병합 처리
	 */
	private void makeMergeTitleRow(Sheet sheet, Class<?> dataClass) {
		
		for (int i = 0; i < super.writeStartRow; i++) {
			sheet.createRow(i);
		}
		
		Row row = sheet.createRow(super.writeStartRow);
		int cellIndex = 0;
		
		boolean needRemoveNewRow = true; 
		
		Field[] fields = dataClass.getDeclaredFields();
		int appendedRows = 0;
		for (Field field : fields) {
			if (field.isAnnotationPresent(Title.class)) {
				Title title = field.getAnnotation(Title.class);
				Merge merge = title.merge();
				
				if (merge.rows() > 1 || merge.cols() > 1) {
					needRemoveNewRow = false;
					
					String titleValue = merge.value();
					if (merge.value().length() == 0) {
						titleValue = title.value();
					}
					
					Cell cell = row.createCell(cellIndex);
					cell.setCellValue(titleValue);
					
					int endRowIndex = row.getRowNum();
					if (merge.rows() - 1 >= 0) {
						endRowIndex += merge.rows() - 1;
					}
					appendedRows = Math.max(appendedRows, endRowIndex);
					
					int endColIndex = cellIndex;
					if (merge.cols() - 1 >= 0) {
						endColIndex += merge.cols() - 1;
					}
					sheet.addMergedRegion(new CellRangeAddress( row.getRowNum(), endRowIndex, cellIndex, endColIndex ));
					
					setCellStyle(cell, field);
					
					cellIndex += merge.cols();
				}
			}
		}
		
		if (!needRemoveNewRow) {
			for (int i = 0; i < appendedRows; i++) {
				sheet.createRow(sheet.getPhysicalNumberOfRows());
			}			
		}
		else {
			sheet.removeRow(row);
		}
		
		this.nextRowIndex = sheet.getPhysicalNumberOfRows();
	}
	
	/**
	 * 타이틀 작성
	 */
	protected void makeMainTitleRow( Sheet sheet, Class<?> dataClass, int startRow ) {
		super.writeStartRow = startRow;
		this.nextRowIndex = 0;
		
		makeMergeTitleRow(sheet, dataClass == null ? super.dataClass : dataClass);
		
		Row row = sheet.createRow(this.nextRowIndex);
		int cellIndex = 0;
		
		Field[] fields = dataClass.getDeclaredFields();
		
		boolean needRemoveNewRow = true;
		for (Field field : fields) {
			if (field.isAnnotationPresent(Title.class)) {
				Title title = field.getAnnotation(Title.class);
				if (title.ignoreTitle()) {
					cellIndex++;
					continue;
				}
				
				Cell cell = null;
				if (title.appendPrevRow()) {
					cell = sheet.getRow(row.getRowNum() - 1).createCell(cellIndex);
				}
				else {
					cell = row.createCell(cellIndex);
					needRemoveNewRow = false;
				}
				
				cell.setCellValue(title.value());
				
				setCellStyle(cell, field);
				cellIndex++;
			}
		}
		
		if (needRemoveNewRow) {
			sheet.removeRow(row);
		}
		
		this.nextRowIndex = sheet.getPhysicalNumberOfRows();
	}
	
	/**
	 * 셀 스타일 적용
	 * @param cell
	 * @param field
	 */
	protected void setCellStyle(Cell cell, Field field) {
		if (super.borderStyle != null) {
			cell.setCellStyle(super.borderStyle);
		}
		if (super.backgroundStyle.containsKey(field.getName())) {
			cell.setCellStyle(super.backgroundStyle.get(field.getName()));
		}
		if (super.textStyle.containsKey(field.getName())) {
			cell.setCellStyle(super.textStyle.get(field.getName()));
		}
		if (super.alignStyle.containsKey(field.getName())) {
			cell.setCellStyle(super.alignStyle.get(field.getName()));
		}
	}
	
}
