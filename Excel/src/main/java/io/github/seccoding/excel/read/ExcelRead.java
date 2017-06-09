package io.github.seccoding.excel.read;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.util.CellRef;
import io.github.seccoding.excel.read.util.FileType;

/**
 * 엑셀 파일을 읽어 온다.
 * <pre>
 * 사용 예제
	ReadOption ro = new ReadOption();
	ro.setFilePath("/Users/mcjang/ktest/uploadedFile/practiceTest.xlsx");
	ro.setOutputColumns("C", "D", "E", "F", "G", "H", "I");
	ro.setStartRow(3);
	
	List&lt;Map&lt;String, String>> result = ExcelRead.read(ro);
	
	for(Map&lt;String, String> map : result) {
		System.out.println(map.get("E"));
	}
</pre>
 * @author Minchang Jang (mc.jang@hucloud.co.kr)
 */
public class ExcelRead {
	
	/**
	 * 엑셀 파일을 읽어옴
	 * @param readOption
	 * @return
	 */
	public static List<Map<String, String>> read(ReadOption readOption) {
		
		Workbook wb = FileType.getWorkbook(readOption.getFilePath());
		Sheet sheet = wb.getSheetAt(0);
		
		int numOfRows = sheet.getPhysicalNumberOfRows();
		int numOfCells = 0;
		
		Row row = null;
		Cell cell = null;
		
		String cellName = "";
		
		Map<String, String> map = null;
		List<Map<String, String>> result = new ArrayList<Map<String, String>>(); 
		
		for(int rowIndex = readOption.getStartRow() - 1; rowIndex < numOfRows; rowIndex++) {
			
			row = sheet.getRow(rowIndex);
			
			if(row != null) {
				numOfCells = row.getPhysicalNumberOfCells();
				
				map = new HashMap<String, String>();
				
				for(int cellIndex = 0; cellIndex < numOfCells; cellIndex++) {
					
					cell = row.getCell(cellIndex);
					cellName = CellRef.getName(cell, cellIndex);
					
					if( !readOption.getOutputColumns().contains(cellName) ) {
						continue;
					}
					
					map.put(cellName, CellRef.getValue(cell));
				}
				
				result.add(map);
				
			}
			
		}
		
		return result;
		
	}
	
}
