package io.github.seccoding.excel;

import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.ExcelRead;

public class ExcelReadTest {

	static ReadOption ro = new ReadOption();

	public static void main(String[] args) {

		ro.setFilePath("Excel File Path");
		ro.setOutputColumns("A", "B", "C", "D", "E", "F");
		ro.setStartRow(1);
		ro.setSheetName("Sheet1");

		test1();
		test2();
		test3();
		test4();
	}

	@Deprecated
	public static void test1() {
		System.out.println("test1");
		Map<String, String> result = new ExcelRead().read(ro);

		System.out.println(result);
		System.out.println(result.get("A3"));
		System.out.println(result.get("B3"));
		System.out.println(result.get("C3"));
	}

	@Deprecated
	public static void test2() {
		System.out.println("test2");
		
		ro.setOutputColumns("A", "B", "C");
		TestClass result = new ExcelRead<TestClass>().readToObject(ro, TestClass.class);

		System.out.println(result.getNo().size());
		System.out.println(result.getColumnName().size());
		System.out.println(result.getType().size());

		System.out.print(result.getNo().get(result.getNo().size() - 1));
		System.out.print(" / " + result.getColumnName().get(result.getColumnName().size() - 1));
		System.out.println(" / " + result.getType().get(result.getType().size() - 1));
	}
	
	public static void test3() {
		System.out.println("test3");
		
		ro.setOutputColumns("A", "B", "C");
		List<TestClass2> result = new ExcelRead<TestClass2>().readToList(ro, TestClass2.class);
		System.out.println(result.size());
		
		for (TestClass2 testClass2 : result) {
			System.out.print(testClass2.getNo());
			System.out.print(" / " + testClass2.getColumnName());
			System.out.println(" / " + testClass2.getType());
		}
		
	}

	public static void test4() {
		System.out.println("test4");
		ro.setOutputColumns("B");
		String result = new ExcelRead().getValue(ro, "B3");
		System.out.println(result);
	}

	@Deprecated
	public static class TestClass {

		@Field("A")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private List<String> no;

		@Field("B")
		private List<String> columnName;

		@Field("C")
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

	public static class TestClass2 {

		@Field("A")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private String no;

		@Field("B")
		private String columnName;

		@Field("C")
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
