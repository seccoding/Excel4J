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
### maven dependency에 Excel4J-2.1.2.jar 파일을 추가할 경우 
<a href="#바로가기">상위로 가기</a>
1. Excel4J-3.0.0.jar파일을 C:\에 복사합니다.
1. Maven 명령어를 이용해 .m2 Repository 에 Excel4J-3.0.0.jar 를 설치(저장)합니다.<pre>mvn install:install-file -Dfile=C:\Excel4J-3.0.0.jar -DgroupId=io.github.seccoding -DartifactId=Excel4J -Dversion=3.0.0 -Dpackaging=jar</pre>
1. 본인의 Project/pom.xml 에 dependency를 추가합니다.<pre>
	&lt;dependency&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;groupId&gt;io.github.seccoding&lt;/groupId&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;artifactId&gt;Excel4J&lt;/artifactId&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;version&gt;3.0.0&lt;/version&gt;
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

```

<a href="#바로가기">상위로 가기</a>
