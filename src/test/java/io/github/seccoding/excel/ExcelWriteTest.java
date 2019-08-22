package io.github.seccoding.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Format;
import io.github.seccoding.excel.annotations.Title;
import io.github.seccoding.excel.option.WriteOption;
import io.github.seccoding.excel.write.ExcelWrite;

/**
 * ExcelWriteTest Example
 * 
 * @author Min Chang Jang (mcjang1116@gmail.com)
 */
public class ExcelWriteTest {

	public static void main(String[] args) {

		WriteOption<TestVO> wo = new WriteOption<TestVO>();
		wo.setFileName("test.xlsx");
		wo.setFilePath("C:\\Users\\mcjan\\Desktop\\");

		List<TestVO> contents = new ArrayList<TestVO>();
		contents.add(new TestVO(111111111, "ABC", true, "=1+1", "2019-02-21"));
		contents.add(new TestVO(2222222, "DEF", true, "=2+2", "2019-02-21"));
		contents.add(new TestVO(33333, "HIJ", true, "=3+3", "2019-02-21"));
		wo.setContents(contents);

		File excelFile = ExcelWrite.write(wo);
	}

	@ExcelSheet(value="TestSheet", useTitle = true)
	public static class TestVO {

		@Title("Title1")
		@Format(alignment = Format.LEFT, verticalAlignment = Format.V_CENTER, bold = true, dataFormat = "#,###")
		private int id;

		@Title("Title2")
		@Format(alignment = Format.LEFT, verticalAlignment = Format.V_CENTER)
		private String content;

		@Title("Title3")
		@Format(alignment = Format.LEFT, verticalAlignment = Format.V_CENTER)
		private boolean isTrue;

		@Title("Title4")
		@Format(alignment = Format.CENTER, verticalAlignment = Format.V_CENTER)
		private String formula;

		@Title(value = "Title5", date = true)
//		@Format(dataFormat = "yyyy-MM-dd") // 2019-02-21
		@Format(dataFormat = "yyyy-MM-dd", toDataFormat="dd-MM-yyyy") // 21-02-2019
		private String date;

		public TestVO(int id, String content, boolean isTrue, String formula, String date) {
			this.id = id;
			this.content = content;
			this.isTrue = isTrue;
			this.formula = formula;
			this.date = date;
		}

		public int getId() {
			return id;
		}

		public void setId(int id) {
			this.id = id;
		}

		public String getContent() {
			return content;
		}

		public void setContent(String content) {
			this.content = content;
		}

		public boolean isTrue() {
			return isTrue;
		}

		public void setTrue(boolean isTrue) {
			this.isTrue = isTrue;
		}

		public String getFormula() {
			return formula;
		}

		public void setFormula(String formula) {
			this.formula = formula;
		}

		public String getDate() {
			return date;
		}

		public void setDate(String date) {
			this.date = date;
		}

	}

}
