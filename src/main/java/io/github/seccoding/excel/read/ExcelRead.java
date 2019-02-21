package io.github.seccoding.excel.read;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellReference;

import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.util.CellReferenceUtil;

public class ExcelRead<T> extends PreparedExcelRead<T> {

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
	
}
