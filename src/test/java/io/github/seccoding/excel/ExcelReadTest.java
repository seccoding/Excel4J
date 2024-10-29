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

	// 엑셀파일의 첫번째 시트에서 두 번째 row부터 읽는다.
	@ExcelSheet(startRow=1)
	public static class TestClass {

		// 엑셀 시트에서 B컬럼 내용만 읽는다.
		@Field("B")
		private String no;

		// 엑셀 시트에서 C컬럼 내용만 읽는다.
		@Field("C")
		private String columnName;

		// 엑셀 시트에서 E컬럼 내용만 읽는다.
		// E컬럼의 타입이 날짜 타입이므로 isDate 적용.
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
