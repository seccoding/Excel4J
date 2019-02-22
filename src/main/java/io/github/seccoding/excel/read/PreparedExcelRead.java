package io.github.seccoding.excel.read;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.util.Add;
import io.github.seccoding.excel.util.CellReferenceUtil;
import io.github.seccoding.excel.util.GetWorkBook;

/**
 * 엑셀 파일을 읽어 온다.
 * <pre>
 * 사용 예제 : 데이터를  Map으로 받아옴.
	ReadOption ro = new ReadOption();
	ro.setFilePath("/Users/mcjang/ktest/uploadedFile/practiceTest.xlsx");
	ro.setOutputColumns("C", "D", "E", "F", "G", "H", "I");
	ro.setStartRow(3);
	
	Map&lt;String, String> result = new ExcelRead&lt;>().read(ro);
	
	for(Map&lt;String, String> map : result) {
		System.out.println(map.get("E1"));
	}
	
	사용 예제 : 데이터를  Class로 받아옴.
	
	ReadOption ro = new ReadOption();
	ro.setFilePath("/Users/mcjang/ktest/uploadedFile/practiceTest.xlsx");
	ro.setOutputColumns("C", "D", "E", "F", "G", "H", "I");
	ro.setStartRow(3);
	
	TestClass result = new ExcelRead&lt;TestClass>().readToObject(ro, TestClass.class);
	
	// src.test.java.io.github.seccoding.excel.ExcelReadTest.java 참고하세요.
</pre>
 * @author Minchang Jang (mcjang1116@gmail.com)
 */
public class PreparedExcelRead<T> {
	
	protected T t;
	
	private Workbook wb;
	protected Sheet sheet;
	
	private String sheetName;
	
	private int numOfRows;
	private int numOfCells;
	
	protected Row row;
	protected Cell cell;
	
	protected String cellName;
	
	protected final List<T> result = new ArrayList<T>();
	protected Class<?> clazz;
	
	protected void makeData(ReadOption readOption, AddData addData, boolean makeList) {
		for(int rowIndex = readOption.getStartRow() - 1; rowIndex <= numOfRows; rowIndex++) {
			
			row = sheet.getRow(rowIndex);
			
			if(row != null) {
				numOfCells = row.getPhysicalNumberOfCells();
				
				for(int cellIndex = 0; cellIndex < numOfCells; cellIndex++) {
					
					cell = row.getCell(cellIndex);
					cellName = CellReferenceUtil.getName(cell, cellIndex);
					if( readOption.isOverOutputColumnIndex(cellName) ) {
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
	
	protected interface AddData {
		public boolean pushData(int rowIndex);
	}
	
	protected boolean addData(int rowNum, String value) {
		java.lang.reflect.Field[] fields = t.getClass().getDeclaredFields();
		boolean isKeepGoing = true;
		for (java.lang.reflect.Field field : fields) {
			if ( isUsedRequireAnnotaion(field) ) {
				Field annotation = field.getAnnotation(Field.class);
				String column = annotation.value();
				if ( column.equalsIgnoreCase(cellName) && (value == null || value.length() == 0) ) {
					return false;
				}
			}
			
			if ( isKeepGoing && isUsedFieldAnnotaion(field) ) {
				Field annotation = field.getAnnotation(Field.class);
				String column = annotation.value();
				if ( column.equalsIgnoreCase(cellName) ) {
					Add.add(field.getName(), t, cellName + rowNum, value);
				}
			}
		}
		
		return isKeepGoing;
	}
	
	protected void setup(String filePath, String sheetName) {
		
		this.sheetName = sheetName;
		
		getWorkbook(filePath);
		
		if ( sheetName == null || sheetName.length() == 0 ) {
			getSheet(0);
		}
		else {
			getSheet(sheetName);
		}
		
		setNumOfRowsAndCells();
	}
	
	protected void setup(ReadOption readOption) {
		getWorkbook(readOption.getFilePath());
		if ( readOption.getSheetName() == null ) {
			extractSheetName();
		}
		if ( readOption.getOutputColumns().isEmpty() ) {
			extractOutputColumns(readOption);
		}
		if ( readOption.getStartRow() <= 0 ) {
			extractStratRow(readOption);
		}
		setNumOfRowsAndCells();
	}
	
	private void getWorkbook(String filePath) {
		wb = GetWorkBook.getWorkbook(filePath);
	}
	
	private void extractSheetName() {
		ExcelSheet sheet = clazz.getAnnotation(ExcelSheet.class);
		
		if ( sheet != null ) {
			sheetName = sheet.value();
		}
		
		if ( sheetName == null || sheetName.length() == 0 ) {
			getSheet(0);
		}
		else {
			getSheet(sheetName);
		}
	}
	
	private void getSheet(int index) {
		sheet = wb.getSheetAt(index);
	}
	
	private void getSheet(String sheetName) {
		sheet = wb.getSheet(sheetName);
	}
	
	private void extractOutputColumns(ReadOption readOption) {
		
		List<String> outputColumns = new ArrayList<String>();
		
		java.lang.reflect.Field[] fields = clazz.getDeclaredFields();
		
		for (java.lang.reflect.Field f : fields) {
			String column = f.getAnnotation(Field.class).value();
			if ( column != null && column.length() > 0 ) {
				outputColumns.add(column);
			}
		}
		
		readOption.setOutputColumns(outputColumns);
	}
	
	private void extractStratRow(ReadOption readOption) {
		
		ExcelSheet sheet = clazz.getAnnotation(ExcelSheet.class);
		
		if ( sheet != null ) {
			readOption.setStartRow( sheet.startRow() );
		}
		
	}
	
	@SuppressWarnings("unchecked")
	protected T createResultInstance() {
		try {
			this.t = (T) this.clazz.newInstance();
			return t;
		} catch (InstantiationException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private void setNumOfRowsAndCells() {
		if ( sheet == null ) {
			throw new RuntimeException("Can not find sheet [" + this.sheetName + "]");
		}
		numOfRows = sheet.getPhysicalNumberOfRows();
		numOfCells = 0;
	}
	
	private boolean isUsedFieldAnnotaion(java.lang.reflect.Field f) {
		return f.getAnnotation(Field.class) != null;
	}
	
	private boolean isUsedRequireAnnotaion(java.lang.reflect.Field f) {
		return f.getAnnotation(Require.class) != null;
	}
	
	
}
