package io.github.seccoding.excel.option;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.CellReference;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;


/**
 * Excel(xls, xlsx) 파일을 읽을 때, 필요한 옵션을 정의한다. 
 * 여기에 정의된 옵션들을 사용해서 실제 파일을 읽어 온다.
 * 
 * @author Minchang Jang (mcjang1116@gmail.com)
 */
public class ReadOption {

	/**
	 * 읽어올 Excel 파일의 위치.
	 */
	private String filePath;
	
	/**
	 * 읽어올 Excel 시트 이름
	 */
	private String sheetName;
	
	/**
	 * Excel에서 읽어올 Column.
	 */
	private List<String> outputColumns;
	
	private List<Short> outputColumnIndex;
	
	/**
	 * Excel에서 추출을 시작하고 싶은 Row.
	 */
	private int startRow;
	
	/**
	 * 읽어올 Excel 파일의 위치를 가져온다.
	 * @return
	 */
	public String getFilePath() {
		return filePath;
	}
	
	/**
	 * 읽어올 Excel 파일의 위치를 지정한다.
	 * @param String filePath : 파일시스템의 경로 (파일명.확장자 포함)
	 */
	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}
	
	/**
	 * 읽어올 Excel 파일의 시트 이름을 가져온다.
	 * @return
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * 읽어올 Excel 파일의 시트 이름을 지정한다.
	 * @param sheetName
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	/**
	 * Excel에서 읽어올 Column을 가져온다.
	 * @return List<String> Excel에서 읽어올 Column
	 */
	public List<String> getOutputColumns() {
		
		if ( this.outputColumns == null ) {
			return new ArrayList<String>();
		}
		
		List<String> temp = new ArrayList<String>();
		temp.addAll(outputColumns);
		
		return temp;
	}
	
	private List<Short> makeAndGetOutputColumnIndex() {
		if ( this.outputColumns == null ) {
			return new ArrayList<Short>();
		}
		
		outputColumnIndex = new ArrayList<Short>();
		
		for (String column : outputColumns) {
			CellReference cr = new CellReference(column + "1");
			outputColumnIndex.add(cr.getCol());
		}
		
		return outputColumnIndex;
		
	}
	
	public boolean isOverOutputColumnIndex(String columnName) {
		columnName = columnName.replaceAll("[0-9]", "");
		CellReference cr = new CellReference(columnName + "1");
		
		short col = cr.getCol();
		
		for(short outputIndex : outputColumnIndex) {
			if ( col <= outputIndex ) {
				return false;
			}
		}
		return true;
	}
	
	/**
	 * Excel에서 읽어올 Column을 지정한다.
	 * @param List<String>
	 */
	public void setOutputColumns(List<String> outputColumns) {
		
		if ( outputColumns == null ) {
			this.outputColumns = null;
			return;
		}
		
		List<String> temp = new ArrayList<String>();
		temp.addAll(outputColumns);
		
		if ( this.outputColumns == null ) {
			this.outputColumns = new ArrayList<String>();
		}
		
		this.outputColumns.clear();
		this.outputColumns = temp;
		
		this.makeAndGetOutputColumnIndex();
	}
	
	/**
	 * Excel에서 읽어올 Column을 지정한다.
	 * @param String[] 가변길이로 지정함.
	 */
	public void setOutputColumns(String outputColumn, String ... outputColumns) {
		
		if ( outputColumns == null ) {
			this.outputColumns = null;
			return;
		}
		
		if(this.outputColumns == null) {
			this.outputColumns = new ArrayList<String>();
		}
		
		this.outputColumns.clear();
		this.outputColumns.add(outputColumn);
		
		for(String ouputColumn : outputColumns) {
			this.outputColumns.add(ouputColumn);
		}
		
		this.makeAndGetOutputColumnIndex();
	}
	
	/**
	 * Excel에서 추출을 시작하고 싶은 Row를 가져온다.
	 * @return int 추출 시작 번호
	 */
	public int getStartRow() {
		return startRow;
	}
	
	/**
	 * Excel에서 추출을 시작하고 싶은 Row를 지정한다.
	 * Excel문서와 동일하게 1부터 시작한다.
	 * @param int 추출 시작 번호
	 */
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	
	public void extractOutputColumns(Class<?> clazz) {
		
		if ( getOutputColumns().isEmpty() ) {
			
			List<String> outputColumns = new ArrayList<String>();
			java.lang.reflect.Field[] fields = clazz.getDeclaredFields();
			
			for (java.lang.reflect.Field f : fields) {
				Field field = f.getAnnotation(Field.class);
				
				if ( field != null ) {
					String column = f.getAnnotation(Field.class).value();
					if ( column != null && column.length() > 0 ) {
						outputColumns.add(column);
					}				
				}
			}
			
			this.setOutputColumns(outputColumns);
		}
	}
	
	public void extractStratRow(Class<?> clazz) {
		if ( getStartRow() <= 0 ) {
			ExcelSheet sheet = clazz.getAnnotation(ExcelSheet.class);
			if ( sheet != null ) {
				this.setStartRow( sheet.startRow() );
			}
		}
	}
	
}
