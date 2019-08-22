package io.github.seccoding.excel.util.read;

import java.util.ArrayList;
import java.util.List;

import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.util.read.share.ReadShare;

public class PreparedExcelRead<T> {
	
	protected T t;
	
	protected String cellName;
	
	protected final List<T> result = new ArrayList<T>();
	
	protected void setup() {
		GetWorkbook.get(ReadShare.readOption.getFilePath());
		GetSheet.getSheetName();
		
		ReadShare.readOption.extractOutputColumns(ReadShare.clazz);
		ReadShare.readOption.extractStratRow(ReadShare.clazz);
	}
	
	protected void setup(String filePath) {
		setup(filePath, null);
	}
	
	protected void setup(String filePath, String sheetName) {
		ReadShare.sheetName = sheetName;
		GetWorkbook.get(filePath);
		GetSheet.set();
	}
	
	protected void makeData(AddData addData, boolean makeList) {
		for(int rowIndex = ReadShare.readOption.getStartRow() - 1; rowIndex <= ReadShare.numOfRows; rowIndex++) {
			GetRow.setRow(rowIndex);
			
			if( GetRow.isNotNull() ) {
				GetRow.setPhysicalNumberOfCells();
				
				for(int cellIndex = 0; cellIndex < ReadShare.numOfCells; cellIndex++) {
					GetCell.setCell(cellIndex);
					cellName = GetCell.getCellName(cellIndex);
					if( ReadShare.readOption.isOverOutputColumnIndex(cellName) ) {
						break;
					}
					
					if ( addData != null ) {
						if ( !addData.pushData(rowIndex) ) {
							return;
						}
					}
					
				}
				
				if ( makeList ) {
					result.add(t);
					createResultInstance();
				}
			}
			
		}
	}
	
	protected boolean addData(int rowNum, String value) {
		java.lang.reflect.Field[] fields = t.getClass().getDeclaredFields();
		boolean isKeepGoing = true;
		for (java.lang.reflect.Field field : fields) {
			if ( isUsedRequireAnnotaion(field) ) {
				Field annotation = field.getAnnotation(Field.class);
				if ( annotation == null ) continue;
				String column = annotation.value();
				if ( column.equalsIgnoreCase(cellName) && (value == null || value.length() == 0) ) {
					return false;
				}
			}
			
			if ( isKeepGoing && isUsedFieldAnnotaion(field) ) {
				Field annotation = field.getAnnotation(Field.class);
				if ( annotation == null ) continue;
				String column = annotation.value();
				if ( column.equalsIgnoreCase(cellName) ) {
					Add.add(field.getName(), t, cellName + rowNum, value);
				}
			}
		}
		
		return isKeepGoing;
	}
	
	
	@SuppressWarnings("unchecked")
	protected T createResultInstance() {
		try {
			this.t = (T) ReadShare.clazz.newInstance();
			return t;
		} catch (InstantiationException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private boolean isUsedFieldAnnotaion(java.lang.reflect.Field f) {
		return f.getAnnotation(Field.class) != null;
	}
	
	private boolean isUsedRequireAnnotaion(java.lang.reflect.Field f) {
		return f.getAnnotation(Require.class) != null;
	}
	
	
}
