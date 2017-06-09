package io.github.seccoding.excel.option;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel(xls, xlsx) 파일을 읽을 때, 필요한 옵션을 정의한다. 
 * 여기에 정의된 옵션들을 사용해서 실제 파일을 읽어 온다.
 * 
 * @author Minchang Jang (mc.jang@hucloud.co.kr)
 */
public class ReadOption {

	/**
	 * 읽어올 Excel 파일의 위치.
	 */
	private String filePath;
	
	/**
	 * Excel에서 읽어올 Column.
	 */
	private List<String> outputColumns;
	
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
	 * Excel에서 읽어올 Column을 가져온다.
	 * @return List<String> Excel에서 읽어올 Column
	 */
	public List<String> getOutputColumns() {
		
		List<String> temp = new ArrayList<String>();
		temp.addAll(outputColumns);
		
		return temp;
	}
	
	/**
	 * Excel에서 읽어올 Column을 지정한다.
	 * @param List<String>
	 */
	public void setOutputColumns(List<String> outputColumns) {
		
		List<String> temp = new ArrayList<String>();
		temp.addAll(outputColumns);
		
		this.outputColumns = temp;
	}
	
	/**
	 * Excel에서 읽어올 Column을 지정한다.
	 * @param String[] 가변길이로 지정함.
	 */
	public void setOutputColumns(String ... outputColumns) {
		
		if(this.outputColumns == null) {
			this.outputColumns = new ArrayList<String>();
		}
		
		for(String ouputColumn : outputColumns) {
			this.outputColumns.add(ouputColumn);
		}
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
	
}
