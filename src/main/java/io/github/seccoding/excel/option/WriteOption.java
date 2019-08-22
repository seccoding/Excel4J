package io.github.seccoding.excel.option;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import io.github.seccoding.excel.annotations.ExcelSheet;


/**
 * Excel(xls, xlsx) 파일을 쓸 때, 필요한 옵션을 정의한다.
 * 여기에 정의된 옵션으로 실제 파일을 작성한다.
 * <pre>
 * 사용 예제
	WriteOption wo = new WriteOption();
	wo.setSheetName("Test");
	wo.setFileName("test.xlsx");
	wo.setFilePath("d:\\");
	List<String> titles = new ArrayList<String>();
	titles.add("Title1");
	titles.add("Title2");
	titles.add("Title3");
	wo.setTitles(titles);
	
	List<String[]> contents = new ArrayList<String[]>();
	contents.add(new String[]{"1", "2", "3"});
	contents.add(new String[]{"11", "22", "33"});
	contents.add(new String[]{"111", "222", "333"});
	wo.setContents(contents);
</pre>
 * @author Minchang Jang (mcjang1116@gmail.com)
 *
 */
public class WriteOption<T> {
	
	/**
	 * 엑셀 파일이 만들어질 위치를 지정한다.
	 */
	private String filePath;
	/**
	 * 엑셀 파일의 이름을 정의한다. 확장자까지 포함해야 한다.
	 */
	private String fileName;
	/**
	 * 엑셀 Document의 Sheet 명을 정의 한다.
	 */
	private String sheetName;
	/**
	 * 엑셀 문서가 타이틀(헤더)을 정의한다.
	 */
	private List<String> titles;
	/**
	 * 엑셀 문서의 내용을 정의한다.
	 */
	private List<T> contents;
	
	/**
	 * 엑셀 파일의 경로를 가져온다.
	 * @return String 엑셀 파일의 절대 경로
	 */
	public String getFilePath() {
		return filePath;
	}
	/**
	 * 엑셀 파일의 절대 경로를 정의한다.
	 * @param String 파일 시스템의 물리적 위치. 예>D:\\tempExcelFolder\\
	 */
	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}
	/**
	 * 엑셀 파일의 이름을 가져온다. 확장자를 포함한다.
	 * @return String 확장자를 포함한 엑셀 파일의 이름
	 */
	public String getFileName() {
		return fileName;
	}
	/**
	 * 엑셀 파일의 이름을 정의한다. 확장자를 포함해야 한다.
	 * @param String 확장자를 포함한 엑셀 파일의 이름
	 */
	public void setFileName(String fileName) {
		
		fileName = fileName.replace("/", "_")
							.replace("\\", "_");
		
		this.fileName = File.separator + fileName;
	}
	/**
	 * 엑셀 문서 내의 Sheet 이름을 가져온다.
	 * @return String 엑셀 문서내의 Sheet 명
	 */
	public String getSheetName() {
		return sheetName;
	}
	/**
	 * 엑셀 문서 내의 Sheet 명을 정의한다. 단 하나의 Sheet만 생성할 수 있다.
	 * @param String 엑셀 문서의 Sheet 이름
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	private void setSheetName() {
		ExcelSheet sheet = contents.get(0).getClass().getAnnotation(ExcelSheet.class);
		this.setSheetName(sheet.value());
	}
	
	/**
	 * 엑셀 문서의 타이틀 정보를 가져온다.
	 * @return List<String> 타이틀 정보를 가진 List
	 */
	public List<String> getTitles() {
		
		List<String> temp = new ArrayList<String>();
		temp.addAll(this.titles);
		
		return temp;
	}
	/**
	 * 엑셀 문서의 타이틀 정보를 정의한다.
	 * @param List<String> List 형태의 타이틀 정보
	 */
	public void setTitles(List<String> titles) {
		
		List<String> temp = new ArrayList<String>();
		temp.addAll(titles);
		
		this.titles = temp;
	}
	/**
	 * 엑셀 문서의 타이틀 정보를 정의한다.
	 * @param String 가변 인자 형태의 타이틀 정보
	 */
	public void setTitles(String ... titles) {
		List<String> temp = Arrays.asList(titles);
		this.titles = temp;
	}
	
	/**
	 * 엑셀 문서에 포함될 내용을 가져온다.
	 * @return List<Stirng[]> 엑셀 문서에 포함될 내용
	 */
	public List<T> getContents() {
		
		List<T> temp = new ArrayList<T>();
		temp.addAll(this.contents);
		
		return temp;
	}
	
	/**
	 * 엑셀 문서의 내용을 정의한다.
	 * @param List<String[]> 리스트 형태의 내용 정보. 하나의 Row는 하나의 배열로 정의한다.
	 */
	public void setContents(List<T> contents) {
		List<T> temp = new ArrayList<T>();
		temp.addAll(contents);
		
		this.contents = temp;
		
		setSheetName();
	}
	
	/**
	 * 엑셀 문서의 내용을 정의한다.
	 * @param String[] 가변 길이 형태. 한번 호출 할 때마다 한 Row씩 증가한다.
	 */
	public void setContents(T ... contents) {
		List<T> temp = new ArrayList<T>();
		for (T t : contents) {
			temp.add(t);
		}
		
		this.contents = temp;
		
		setSheetName();
	}
	
}
