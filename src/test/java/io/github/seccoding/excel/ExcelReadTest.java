package io.github.seccoding.excel;

import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.ExcelRead;

public class ExcelReadTest {

	private static ReadOption ro = new ReadOption();
	private static String filePath = "C:\\Users\\mcjan\\Desktop\\ktds Table Descriptor.xlsx";
	
	public static void main(String[] args) {

		ro.setFilePath(filePath);
		ro.setOutputColumns("B", "C", "F");
		ro.setStartRow(6);
//		ro.setSheetName("Sheet1");
		
		test1();
		test2();
		test3();
		test4();
		test5();
		test6();
	}

	@Deprecated
	public static void test1() {
		System.out.println("test1");
		ro.setSheetName("ADMIN");
		Map<String, String> result = new ExcelRead().read(ro);

		System.out.println(result);
		System.out.println(result.get("B7"));
		System.out.println(result.get("C7"));
		System.out.println(result.get("F7"));
	}

	@Deprecated
	public static void test2() {
		System.out.println("test2");
		
		ro.setOutputColumns(null); // TestClass의 @Field로 대체
		ro.setSheetName(null); // TestClass의 @ExcelSheet() 로 대체
		
		TestClass result = new ExcelRead<TestClass>().readToObject(ro, TestClass.class);

		System.out.println(result.getNo().size());
		System.out.println(result.getColumnName().size());
		System.out.println(result.getType().size());
	}
	
	public static void test3() {
		System.out.println("test3");
		
		ro.setOutputColumns(null); // TestClass의 @Field로 대체
		ro.setSheetName(null); // TestClass의 @ExcelSheet() 로 대체
		
		List<TestClass2> result = new ExcelRead<TestClass2>().readToList(ro, TestClass2.class);
		System.out.println(result.size());
	}

	@Deprecated
	public static void test4() {
		System.out.println("test4");
		ro.setSheetName("ADMIN");
		String result = new ExcelRead().getValue(ro, "B3");
		System.out.println(result);
	}
	
	public static void test5() {
		System.out.println("test5");
		String result = new ExcelRead().getValue(filePath, "ADMIN", "B3");
		System.out.println(result);
	}
	
	public static void test6() {
		System.out.println("test6");
		String result = new ExcelRead().getValue(filePath, "B3");
		System.out.println(result);
	}

	@ExcelSheet("ADMIN")
	@Deprecated
	public static class TestClass {

		@Field("B")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private List<String> no;

		@Field("C")
		private List<String> columnName;

		@Field("F")
		private List<String> type;

		public List<String> getNo() {
			return no;
		}

		public void setNo(List<String> no) {
			this.no = no;
		}

		public List<String> getColumnName() {
			return columnName;
		}

		public void setColumnName(List<String> columnName) {
			this.columnName = columnName;
		}

		public List<String> getType() {
			return type;
		}

		public void setType(List<String> type) {
			this.type = type;
		}

	}

	@ExcelSheet("ADMIN")
	public static class TestClass2 {

		@Field("H")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private String no;

		@Field("J")
		private String columnName;

		@Field("M")
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
