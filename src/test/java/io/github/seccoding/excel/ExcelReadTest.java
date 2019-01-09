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

		ro.setFilePath("C:\\Users\\mcjan\\Desktop\\ktds Table Descriptor.xlsx");
		ro.setOutputColumns("B", "C", "F");
		ro.setStartRow(7);
		ro.setSheetName("ADMIN");
		
		
		test1();
		test2();
	}

	public static void test1() {
		Map<String, String> result = new ExcelRead().read(ro);

		System.out.println(result);
	}

	public static void test2() {
		TestClass result = new ExcelRead<TestClass>().readToObject(ro, TestClass.class);

		System.out.println(result.getNo());
		System.out.println(result.getColumnName());
		System.out.println(result.getType());
	}

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

}
