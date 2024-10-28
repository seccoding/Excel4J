package io.github.seccoding.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import io.github.seccoding.excel.annotations.Align;
import io.github.seccoding.excel.annotations.BackgroundColor;
import io.github.seccoding.excel.annotations.Border;
import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Text;
import io.github.seccoding.excel.annotations.Title;
import io.github.seccoding.excel.write.Write;

/**
 * ExcelWriteTest Example
 * 
 * @author Min Chang Jang (mcjang1116@gmail.com)
 */
public class ExcelWriteTest {

	public static void main(String[] args) {
		List<TestVO> contents = new ArrayList<TestVO>();
		contents.add(new TestVO(111111111, "ABC", true, "=1+1", "2019-02-21"));
		contents.add(new TestVO(2222222, "DEF", true, "=2+2", "2019-02-21"));
		contents.add(new TestVO(33333, "HIJ", true, "=3+3", "2019-02-21"));

		Write<TestVO> write = new Write<>(TestVO.class, contents);
		write.write(new File("/Users/codemakers/Desktop", "Test.xlsx"));
	}

	@ExcelSheet(value = "TestSheet")
	@Border(value = BorderStyle.MEDIUM, color = IndexedColors.RED)
	public static class TestVO {

		@Title(value="Title1")
		@BackgroundColor(IndexedColors.BLACK)
		@Text(color = IndexedColors.WHITE, bold = true) 
		@Align(value=HorizontalAlignment.CENTER, verticalAlignment = VerticalAlignment.TOP)
		private int id;

		@Title(value="Title2")
		@BackgroundColor(IndexedColors.WHITE)
		@Text(color = IndexedColors.RED)
		@Align(value=HorizontalAlignment.RIGHT, verticalAlignment = VerticalAlignment.CENTER)
		private String content;

		@Title(value="Title3")
		@BackgroundColor(IndexedColors.RED)
		@Text(color = IndexedColors.YELLOW)
		@Align(value=HorizontalAlignment.LEFT, verticalAlignment = VerticalAlignment.BOTTOM)
		private boolean isTrue;

		@Title(value="Title4")
		@BackgroundColor(IndexedColors.BLUE)
		@Text(color = IndexedColors.BLUE_GREY)
		private String formula;

		@Title(value="Title5")
		@BackgroundColor(IndexedColors.YELLOW)
		@Text(color = IndexedColors.BROWN, bold = true)
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

		public boolean getIsTrue() {
			return isTrue;
		}

		public void setIsTrue(boolean isTrue) {
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
