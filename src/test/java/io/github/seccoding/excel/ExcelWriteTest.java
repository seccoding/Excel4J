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
		List<TestVO> contents = new ArrayList<>();
		contents.add(new TestVO(111111111, "ABC", true, "=1+1", "2019-02-21"));
		contents.add(new TestVO(2222222, "DEF", true, "=2+2", "2019-02-21"));
		contents.add(new TestVO(33333, "HIJ", true, "=3+3", "2019-02-21"));
		
		Write<TestVO> write = new Write<>(TestVO.class, contents);
		write.write("Test.xlsx");

		
		List<TestVO2> contents2 = new ArrayList<>();
		contents2.add(new TestVO2(5555, "ㅁㅁㅁABC", true));
		contents2.add(new TestVO2(6666, "ㅇㅇㅇDEF", true));
		contents2.add(new TestVO2(7777, "ㅗㅗㅗHIJ", true));
		
		write.appendNewSheet(TestVO2.class, contents2);
		write.toFile(new File("/Users/codemakers/Desktop", "Test.xlsx"));
	}

	// 엑셀 파일의 "TestSheet" 시트에 내용을 작성한다.
	@ExcelSheet(value = "TestSheet")
	@Border(value = BorderStyle.MEDIUM, color = IndexedColors.RED)
	public static class TestVO {

		// 첫 번째 컬럼의 타이틀(헤더)을 "Title1" 로 작성한다.
		@Title(value = "Title1")
		// 해당 컬럼의 배경색을 검은색으로 지정한다.
		@BackgroundColor(IndexedColors.BLACK)
		// 해당 컬럼의 글자를 굵은 흰색으로 지정한다.
		@Text(color = IndexedColors.WHITE, bold = true)
		// 해당 컬럼은 가로(중앙), 세로(위)로 정렬한다.
		@Align(value = HorizontalAlignment.CENTER, verticalAlignment = VerticalAlignment.TOP)
		private int id;

		// 첫 번째 컬럼의 타이틀(헤더)을 "Title2" 로 작성한다.
		@Title(value = "Title2")
		// 해당 컬럼의 배경색을 하양색으로 지정한다.
		@BackgroundColor(IndexedColors.WHITE)
		// 해당 컬럼의 글자를 붉은색으로 지정한다.
		@Text(color = IndexedColors.RED)
		// 해당 컬럼은 가로(오른쪽), 세로(중앙)로 정렬한다.
		@Align(value = HorizontalAlignment.RIGHT, verticalAlignment = VerticalAlignment.CENTER)
		private String content;

		// 첫 번째 컬럼의 타이틀(헤더)을 "Title3" 로 작성한다.
		@Title(value = "Title3")
		// 해당 컬럼의 배경색을 붉은색으로 지정한다.
		@BackgroundColor(IndexedColors.RED)
		// 해당 컬럼의 글자를 노랑색으로 지정한다.
		@Text(color = IndexedColors.YELLOW)
		// 해당 컬럼은 가로(왼쪽), 세로(아래)로 정렬한다.
		@Align(value = HorizontalAlignment.LEFT, verticalAlignment = VerticalAlignment.BOTTOM)
		private boolean isTrue;

		// 첫 번째 컬럼의 타이틀(헤더)을 "Title4" 로 작성한다.
		@Title(value = "Title4")
		// 해당 컬럼의 배경색을 푸른색으로 지정한다.
		@BackgroundColor(IndexedColors.BLUE)
		// 해당 컬럼의 글자를 BlueGrey색으로 지정한다.
		@Text(color = IndexedColors.BLUE_GREY)
		private String formula;

		// 첫 번째 컬럼의 타이틀(헤더)을 "Title5" 로 작성한다.
		@Title(value = "Title5")
		// 해당 컬럼의 배경색을 노랑색으로 지정한다.
		@BackgroundColor(IndexedColors.YELLOW)
		// 해당 컬럼의 글자를 굵은 갈색으로 지정한다.
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

	@ExcelSheet(value = "TestSheet2222")
	public static class TestVO2 {

		@Title(value = "_Title1_")
		private int id;

		@Title(value = "_Title2_")
		private String content;

		@Title(value = "_Title3_")
		private boolean isTrue;

		public TestVO2(int id, String content, boolean isTrue) {
			this.id = id;
			this.content = content;
			this.isTrue = isTrue;
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

	}

}
