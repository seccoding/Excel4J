package io.github.seccoding.excel.read;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.util.CellRef;
import io.github.seccoding.excel.read.util.FileType;

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
 * @author Minchang Jang (mc.jang@hucloud.co.kr)
 */
public class ExcelRead<T> {
	
	private static final int LIST = 1;
	private static final int MAP = 2;
	private static final int SET = 3;
	
	private T t;
	
	private Workbook wb;
	private Sheet sheet;
	
	private int numOfRows;
	private int numOfCells;
	
	private Row row;
	private Cell cell;
	
	private String cellName;
	
	
	/**
	 * 엑셀 파일을 읽어옴
	 * @param readOption
	 * @return
	 */
	public Map<String, String> read(ReadOption readOption) {
		
		setup(readOption.getFilePath(), readOption.getSheetName());
		final Map<String, String> map = new HashMap<String, String>();
		
		makeData(readOption, new AddData() {
			@Override
			public void pushData(int rowIndex) {
				map.put(cellName + (rowIndex + 1), CellRef.getValue(cell));
				
			}});
		return map;
		
	}
	
	/**
	 * 엑셀 파일을 읽어옴
	 * @param readOption
	 * @return
	 */
	public T readToObject(ReadOption readOption, Class<?> claz) {
		
		setup(readOption.getFilePath(), readOption.getSheetName());
		createResultInstance(claz);
		
		makeData(readOption, new AddData() {
			@Override
			public void pushData(int rowIndex) {
				boolean isKeepGoing = addData(t, cellName, rowIndex + 1, CellRef.getValue(cell));
				if ( !isKeepGoing ) {
					return;
				}
			}
		});
		
		return t;
		
	}
	
	public String getValue(ReadOption readOption, String cellName) {
		setup(readOption.getFilePath(), readOption.getSheetName());
		
		CellReference cr = new CellReference(cellName);
		row = sheet.getRow(cr.getRow());
		cell = row.getCell(cr.getCol());
		
		return CellRef.getValue(cell);
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
	
	private void getWorkbook(String filePath) {
		wb = FileType.getWorkbook(filePath);
	}
	
	private void getSheet(int index) {
		sheet = wb.getSheetAt(index);
	}
	
	private void getSheet(String sheetName) {
		sheet = wb.getSheet(sheetName);
	}
	
	private void setNumOfRowsAndCells() {
		numOfRows = sheet.getPhysicalNumberOfRows();
		numOfCells = 0;
	}
	
	private T createResultInstance(Class<?> claz) {
		try {
			this.t = (T) claz.newInstance();
			return t;
		} catch (InstantiationException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private interface AddData {
		public void pushData(int rowIndex);
	}
	
	private void makeData(ReadOption readOption, AddData addData) {
		for(int rowIndex = readOption.getStartRow() - 1; rowIndex < numOfRows; rowIndex++) {
			
			row = sheet.getRow(rowIndex);
			
			if(row != null) {
				numOfCells = row.getPhysicalNumberOfCells();
				
				for(int cellIndex = 0; cellIndex < numOfCells; cellIndex++) {
					
					cell = row.getCell(cellIndex);
					cellName = CellRef.getName(cell, cellIndex);
					
					if( !readOption.getOutputColumns().contains(cellName) ) {
						continue;
					}
					
					if ( addData != null ) {
						addData.pushData(rowIndex);
					}
					
				}
				
			}
			
		}
	}
	
	private boolean addData(T t, String columnName, int rowNum, String value) {
		Field[] fields = t.getClass().getDeclaredFields();
			
		for (Field field : fields) {
			if ( field.isAnnotationPresent(Require.class) ) {
				if ( value == null || value.length() == 0 ) {
					return false;
				}
			}
			if ( field.isAnnotationPresent(io.github.seccoding.excel.annotations.Field.class) ) {
				io.github.seccoding.excel.annotations.Field annotation = 
						field.getAnnotation(io.github.seccoding.excel.annotations.Field.class);
				
				String column = annotation.value();
				if ( column.equalsIgnoreCase(columnName) ) {
					int collectionType = getCollectionType(field, t);
					if ( collectionType == ExcelRead.LIST ) {
						List list = getList(field, t);
						list.add(value);
					}
					else if ( collectionType == ExcelRead.MAP ) {
						Map map = getMap(field, t);
						map.put(columnName + rowNum, value);
					}
					else if ( collectionType == ExcelRead.SET ) {
						Set set = getSet(field, t);
						set.add(value);
					}
				}
			}
			
		}
		
		return true;
	}
	
	private int getCollectionType(Field f, T t) {
		
		if ( f.getType() == List.class ) {
			return ExcelRead.LIST;
		}
		else if ( f.getType() == Map.class ) {
			return ExcelRead.MAP;
		}
		else if ( f.getType() == Set.class ) {
			return ExcelRead.SET;
		}
		
		return -1;
	}
	
	private Object getField(Field f, T t) {
		f.setAccessible(true);
		try {
			return f.get(t);
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private List getList(Field f, T t) {
		
		List result = (List) getField(f, t);
		if ( result == null ) {
			result = new ArrayList();
			set(f, result);
		}
		
		return result;
	}
	
	private Map getMap(Field f, T t) {
		
		Map result = (Map) getField(f, t);
		if ( result == null ) {
			result = new HashMap();
			set(f, result);
		}
		
		return result;
	}
	
	private Set getSet(Field f, T t) {
		
		Set result = (Set) getField(f, t);
		if ( result == null ) {
			result = new HashSet();
			set(f, result);
		}
		
		return result;
	}
	
	public void set(Field f, Object obj) {
		try {
			f.set(t, obj);
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
}
