package io.github.seccoding.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Format;
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
//		wo.setSheetName("Test"); @ExcelSheet()로 교체
		wo.setFileName("test.xlsx");
		wo.setFilePath("C:\\Users\\mcjan\\Desktop");

		List<String> titles = new ArrayList<String>();
		titles.add("Title1");
		titles.add("Title2");
		titles.add("Title3");
		titles.add("Title4");
		wo.setTitles(titles);

		List<TestVO> contents = new ArrayList<TestVO>();
		contents.add(new TestVO(1, "ABC", true, "=1+1"));
		contents.add(new TestVO(2, "DEF", true, "=2+2"));
		contents.add(new TestVO(3, "HIJ", true, "=3+3"));
		wo.setContents(contents);

		File excelFile = ExcelWrite.write(wo);
	}

	@ExcelSheet("TestSheet")
	public static class TestVO {
		
		@Field("Title1")
		@Format(alignment=Format.LEFT, verticalAlignment=Format.V_TOP)
		private int id;
		
		@Field("Title2")
		@Format(alignment=Format.LEFT, verticalAlignment=Format.V_TOP)
		private String content;
		
		@Field("Title3")
		@Format(alignment=Format.LEFT, verticalAlignment=Format.V_TOP)
		private boolean isTrue;
		
		@Field("Title4")
		@Format(alignment=Format.CENTER, verticalAlignment=Format.V_TOP)
		private String formula;

		public TestVO(int id, String content, boolean isTrue, String formula) {
			this.id = id;
			this.content = content;
			this.isTrue = isTrue;
			this.formula = formula;
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

	}

}
