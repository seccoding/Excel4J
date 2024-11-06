# Excel4J
Java 에서 엑셀파일을 읽고 쓰는 유틸리티<br/>
xls 와 xlsx를 모두 지원함.

## 바로가기
<a href="#release-note">Release Note</a><br/>
<a href="#사용-방법">사용 방법</a><br/>
<a href="#excel-file-읽기">Excel File 읽기</a><br/>
<a href="#excel-file-쓰기">Excel File 쓰기</a>


## Release Note 
<a href="#바로가기">상위로 가기</a>

### 3.0.2 (2024.11.06) 
<a href="#바로가기">상위로 가기</a>
> - io.github.seccoding.excel.write.Write.appendNewSheet(String sheetName, List<T> data): void
> - io.github.seccoding.excel.write.Write.appendNewSheet(Class<?> dataClass, List<? extends Object> data): void
> - io.github.seccoding.excel.write.Write.appendNewSheet(String sheetName, Class<?> dataClass, List<? extends Object> data): void
> - io.github.seccoding.excel.write.Write.appendNewSheet(int writeStartRow, Class<?> dataClass, List<? extends Object> data): void
> - io.github.seccoding.excel.write.Write.appendNewSheet(String sheetName, int writeStartRow, Class<?> dataClass, List<? extends Object> data): void
> - 추가

### 3.0.1 (2024.11.06) 
<a href="#바로가기">상위로 가기</a>
> - io.github.seccoding.excel.read.Read.read(String sheetName, int startRow): List<T>
> - io.github.seccoding.excel.read.Read.readToMap(): Map<String, List<T>>
> - io.github.seccoding.excel.read.Read.readToMap(Map<String, Integer> sheetsMap): Map<String, List<T>>
> - 추가
 
### 3.0.0 (2024.10.28) 
<a href="#바로가기">상위로 가기</a>
> 전체 구조 변경
> - static 제거.
> - WriteOption, ReadOption 제거.
> - bintray 서비스 종료로 DependencyRepository 제거.

### 2.1.2 (2019.02.22) 
<a href="#바로가기">상위로 가기</a>
> Deprecated
> - ExcelRead.getValue(ReadOption readOption, String cellName)
> 
> Make New
> - ExcelRead.getValue(String filePath, String cellName): String
> - ExcelRead.getValue(String filePath, String sheetName, String cellName): String
> - @ExcelSheet Annotation
> - @Format Annotation
> 
> Modify
> - WriteOption.setSheetName() 대신 @ExcelSheet 로 대체
> - ReadOption.setSheetName() 대신 @ExcelSheet로 대체
> - ReadOption.setOutputColumns() 대신 @Field로 대체
> - ReadOption.setStartRow() 대신 @ExcelSheet로 대체

### 2.1.0 (2019.02.21) 
<a href="#바로가기">상위로 가기</a>
> 1. ExcelRead.read(ReadOption readOption):Map<String, String> is deprecated.
> 2. ExcelRead.readToObject(ReadOption readOption, Class<?> clazz):T is deprecated
> 3. Make new ExcelRead.readToList(ReadOption readOption, Class<?> clazz):List<T>

### 2.0.0 (2019.02.20) 
<a href="#바로가기">상위로 가기</a>
> 1. WriteOption.setContents(List<String[]> contents); 삭제.
> 2. WriteOption<T>.setContents(List<T> contents); 추가
> 2-1. String[] 대신 Data Class 로 사용함.

## 사용 방법
<a href="#바로가기">상위로 가기</a>
### maven dependency에 Excel4J-3.0.2.jar 파일을 추가할 경우 
<a href="#바로가기">상위로 가기</a>
1. Excel4J-3.0.2.jar파일을 C:\에 복사합니다.
1. Maven 명령어를 이용해 .m2 Repository 에 Excel4J-3.0.2.jar 를 설치(저장)합니다.<pre>mvn install:install-file -Dfile=C:\Excel4J-3.0.2.jar -DgroupId=io.github.seccoding -DartifactId=Excel4J -Dversion=3.0.2 -Dpackaging=jar</pre>
1. 본인의 Project/pom.xml 에 dependency를 추가합니다.<pre>
	&lt;dependency&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;groupId&gt;io.github.seccoding&lt;/groupId&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;artifactId&gt;Excel4J&lt;/artifactId&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;version&gt;3.0.2&lt;/version&gt;
	&lt;/dependency&gt;
</pre>

### 소스코드를 사용할 경우 
<a href="#바로가기">상위로 가기</a>
1. Clone or Download 를 클릭합니다.
1. Download ZIP 을 클릭해 소스코드를 다운로드 받습니다.
1. Excel4J/pom.xml의 dependencies를 본인의 Project/pom.xml 에 붙혀넣습니다.
1. Excel4J/src 이하의 자바코드를 본인의 Project에 붙혀넣습니다. 
---
## Excel File 읽기 
<a href="#바로가기">상위로 가기</a>
```java
import java.io.File;
import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.read.Read;

public class ExcelReadTest {

	public static void main(String[] args) {
		File file = new File("/Users/codemakers/Desktop", "Test.xlsx");
		
		Read<TestClass> read = new Read<>(file.toPath(), TestClass.class);
		List<TestClass> result1 = read.read();
		result1.forEach(tc -> {
			System.out.println(tc.getColumnName());
			System.out.println(tc.getNo());
			System.out.println(tc.getType());
		});
		
		List<TestClass> result2 = read.read("testsheet", 3);
		result2.forEach(tc -> {
			System.out.println(tc.getColumnName());
			System.out.println(tc.getNo());
			System.out.println(tc.getType());
		});
		
		
		Map<String, List<TestClass>> resultMap = read.readToMap();
		resultMap.entrySet().forEach(tc -> {
			System.out.println(tc.getKey());
			System.out.println(tc.getValue());
		});
		
		Map<String, List<TestClass>> resultMap2 = read.readToMap(Map.of("testsheet", 1));
		resultMap2.entrySet().forEach(tc -> {
			System.out.println(tc.getKey());
			System.out.println(tc.getValue());
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
```

---
## Excel File 쓰기 
<a href="#바로가기">상위로 가기</a>
```java
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

```

<a href="#바로가기">상위로 가기</a>
