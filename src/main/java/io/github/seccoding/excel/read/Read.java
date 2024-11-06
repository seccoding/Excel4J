package io.github.seccoding.excel.read;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;

import io.github.seccoding.excel.read.abstracts.ReadWorkbook;

/**
 * 엑셀 워크시트의 내용을 읽음.
 * @param <T>
 */
public class Read<T> extends ReadWorkbook<T> {

	/**
	 * @param excelFilePath 읽으려는 엑셀파일의 경로
	 * @param resultClass 엑셀 내용을 담으려는 클래스.
	 */
	public Read(Path excelFilePath, Class<T> resultClass) {
		super(resultClass);
		super.loadWorkbook(excelFilePath);
	}
	
	/**
	 * 엑셀 파일의 내용을 읽어냄.
	 * @return
	 */
	public List<T> read() {
		return super.setValueInExcel();
	}
	
	/**
	 * 엑셀 파일의 특정 시트의 내용을 읽어냄.
	 * @param sheetName 읽을 시트 명 (대소문자 구분 없음)
	 * @param startRow 시트에서 읽을 시작 행 번호
	 * @return
	 */
	public List<T> read(String sheetName, int startRow) {
		Sheet findOneSheet =  super.getAllSheets()
					.stream()
					.filter(sheet -> sheet.getSheetName().equalsIgnoreCase(sheetName))
					.findFirst()
					.orElse(null);
		
		if (findOneSheet == null) {
			return new ArrayList<>();
		}
		
		return super.setValueInExcel( findOneSheet, startRow );
	}
	
	/**
	 * 엑셀 파일의 모든 시트를 읽어냄
	 * @return Map&lt;SheetName: String, Datas: List&lt;T&gt;&gt;
	 */
	public Map<String, List<T>> readToMap() {
		Map<String, List<T>> sheetsData = new HashMap<>();
		
		super.getAllSheets().forEach(sheet -> {
			List<T> rows = super.setValueInExcel(sheet, 0);
			sheetsData.put(sheet.getSheetName(), rows);
		});
		
		return sheetsData;
	}
	
	/**
	 * 엑셀 파일에서 지정한 시트를 읽어냄.
	 * @param sheetsMap ( key: 시트명, value: 읽기 시작할 행 번호 )
	 * @return Map&lt;SheetName: String, Datas: List&lt;T&gt;&gt;
	 */
	public Map<String, List<T>> readToMap(Map<String, Integer> sheetsMap) {
		Map<String, List<T>> sheetsData = new HashMap<>();
		
		sheetsMap.entrySet().stream().forEach(entry -> {
			Sheet readTargetSheet = super.getAllSheets()
					.stream()
					.filter(sheet -> sheet.getSheetName().equalsIgnoreCase(entry.getKey()))
					.findFirst()
					.orElse(null);
			
			if (readTargetSheet != null) {
				List<T> rows = super.setValueInExcel(readTargetSheet, entry.getValue());
				sheetsData.put(entry.getKey(), rows);
			}
		});
		return sheetsData;
	}
	
}
