package io.github.seccoding.excel.read;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

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
public class ExcelRead<T> {
	
	private T t;
	
	private Workbook wb;
	private Sheet sheet;
	
	private int numOfRows;
	private int numOfCells;
	
	private Row row;
	private Cell cell;
	
	private String cellName;
	
	private final List<T> result = new ArrayList<T>();
	private Class<?> clazz;
	
	/**
	 * 엑셀 파일을 읽어옴
	 * @param readOption
	 * @return
	 */
	@Deprecated
	public Map<String, String> read(ReadOption readOption) {
		
		setup(readOption.getFilePath(), readOption.getSheetName());
		final Map<String, String> map = new HashMap<String, String>();
		
		makeData(readOption, new AddData() {
			@Override
			public boolean pushData(int rowIndex) {
				map.put(cellName + (rowIndex + 1), CellReferenceUtil.getValue(cell));
				return true;
			}
		}, false);
		return map;
		
	}
	
	/**
	 * 엑셀 파일을 읽어옴
	 * @param readOption
	 * @return
	 */
	@Deprecated
	public T readToObject(ReadOption readOption, Class<?> clazz) {
		this.clazz = clazz;
		
		setup(readOption);
		createResultInstance();
		
		makeData(readOption, new AddData() {
			@Override
			public boolean pushData(int rowIndex) {
				return addData(rowIndex + 1, CellReferenceUtil.getValue(cell));
			}
		}, false);
		
		return t;
		
	}
	
	/**
	 * 엑셀 파일을 읽어옴
	 * @param readOption
	 * @return
	 */
	public List<T> readToList(ReadOption readOption, Class<?> clazz) {
		this.clazz = clazz;
		
		setup(readOption);
		createResultInstance();
		
		makeData(readOption, new AddData() {
			@Override
			public boolean pushData(int rowIndex) {
				return addData(rowIndex + 1, CellReferenceUtil.getValue(cell));
			}
		}, true);
		
		return result;
		
	}
	
	private void makeData(ReadOption readOption, AddData addData, boolean makeList) {
		for(int rowIndex = readOption.getStartRow() - 1; rowIndex <= numOfRows; rowIndex++) {
			
			row = sheet.getRow(rowIndex);
			
			if(row != null) {
				numOfCells = row.getPhysicalNumberOfCells();
				
				for(int cellIndex = 0; cellIndex < numOfCells; cellIndex++) {
					
					cell = row.getCell(cellIndex);
					cellName = CellReferenceUtil.getName(cell, cellIndex);
					
					if( !readOption.getOutputColumns().contains(cellName) ) {
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
	
	private interface AddData {
		public boolean pushData(int rowIndex);
	}
	
	private boolean addData(int rowNum, String value) {
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
	
	@Deprecated
	public String getValue(ReadOption readOption, String cellName) {
		setup(readOption.getFilePath(), readOption.getSheetName());
		
		CellReference cr = new CellReference(cellName);
		row = sheet.getRow(cr.getRow());
		cell = row.getCell(cr.getCol());
		
		return CellReferenceUtil.getValue(cell);
	}
	
	public String getValue(String filePath, String cellName) {
		setup(filePath, null);
		
		CellReference cr = new CellReference(cellName);
		row = sheet.getRow(cr.getRow());
		cell = row.getCell(cr.getCol());
		
		return CellReferenceUtil.getValue(cell);
	}
	
	public String getValue(String filePath, String sheetName, String cellName) {
		setup(filePath, sheetName);
		
		CellReference cr = new CellReference(cellName);
		row = sheet.getRow(cr.getRow());
		cell = row.getCell(cr.getCol());
		
		return CellReferenceUtil.getValue(cell);
	}
	
	private void setup(String filePath, String sheetName) {
		getWorkbook(filePath);
		
		if ( sheetName == null || sheetName.length() == 0 ) {
			getSheet(0);
		}
		else {
			getSheet(sheetName);
		}
		
		setNumOfRowsAndCells();
	}
	
	private void setup(ReadOption readOption) {
		getWorkbook(readOption.getFilePath());
		if ( readOption.getSheetName() == null ) {
			extractSheetName();
		}
		if ( readOption.getOutputColumns() == null ) {
			extractOutputColumns(readOption);
		}
		setNumOfRowsAndCells();
	}
	
	private void getWorkbook(String filePath) {
		wb = GetWorkBook.getWorkbook(filePath);
	}
	
	private void extractSheetName() {
		ExcelSheet sheet = clazz.getAnnotation(ExcelSheet.class);
		
		String sheetName = null;
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
	
	@SuppressWarnings("unchecked")
	private T createResultInstance() {
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
