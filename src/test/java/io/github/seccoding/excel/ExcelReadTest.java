package io.github.seccoding.excel;

import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.ExcelRead;

public class ExcelReadTest {

	private static String filePath = "Excel File Path";
	
	public static void main(String[] args) {

		testUseReadOption();
		testGetValueInSheet();
		testGetValueInDefaultSheet();
	}

	public static void testUseReadOption() {
		System.out.println("testUseReadOption");
		ReadOption ro = new ReadOption();
		ro.setFilePath(filePath);
		
		List<TestClass> result = new ExcelRead<TestClass>().readToList(ro, TestClass.class);
		System.out.println(result.size());
	}
	
	public static void testUnuseReadOption() {
		System.out.println("testUnuseReadOption");
		List<TestClass> result = new ExcelRead<TestClass>().readToList(filePath, TestClass.class);
		System.out.println(result.size());
	}
	
	public static void testGetValueInSheet() {
		System.out.println("testGetValueInSheet");
		String result = new ExcelRead<>().getValue(filePath, "Sheet1", "B3");
		System.out.println(result);
	}
	
	public static void testGetValueInDefaultSheet() {
		System.out.println("testGetValueInDefaultSheet");
		String result = new ExcelRead<>().getValue(filePath, "B3");
		System.out.println(result);
	}

	@ExcelSheet(value="Sheet1", startRow=1)
	public static class TestClass {

		@Field("B")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private String no;

		@Field("C")
		private String columnName;

		@Field("D")
		private String type;

		public String getNo() {
			return no;
		}

		public void setNo(String no) {
			this.no = no;
		}

		public String getColumnName() {
			return columnName;
		}

		public void setColumnName(String columnName) {
			this.columnName = columnName;
		}

		public String getType() {
			return type;
		}

		public void setType(String type) {
			this.type = type;
		}

	}

}
