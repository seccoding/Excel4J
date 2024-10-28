package io.github.seccoding.excel;

import java.io.File;
import java.util.List;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.read.Read;

public class ExcelReadTest {

	public static void main(String[] args) {
		File file = new File("/Users/codemakers/Desktop", "Test.xlsx");
		
		Read<TestClass> read = new Read<>(file.toPath(), TestClass.class);
		List<TestClass> result = read.read();
		
		result.forEach(tc -> {
			System.out.println(tc.getColumnName());
			System.out.println(tc.getNo());
			System.out.println(tc.getType());
		});
		
	}

	@ExcelSheet(startRow=1)
	public static class TestClass {

		@Field("B")
		private String no;

		@Field("C")
		private String columnName;

		@Field(value="E", isDate = true)
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
