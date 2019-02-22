[ ![Download](https://api.bintray.com/packages/mcjang1116/io.github.seccoding/Excel4J/images/download.svg?version=2.1.1) ](https://bintray.com/mcjang1116/io.github.seccoding/Excel4J/2.1.1/link)

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
### 2.1.1 (2019.02.22) 
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
### maven 사용 
<a href="#바로가기">상위로 가기</a>
1. Repository 추가<pre>
   &lt;repositories&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&lt;repository&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;id&gt;bintray&lt;/id&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;url&gt;http://jcenter.bintray.com</url&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;snapshots&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;enabled&gt;false&lt;/enabled&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&lt;/snapshots&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&lt;/repository&gt;
&lt;/repositories&gt;
   </pre>
   
1. dependency 추가<pre>
   &lt;dependency&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&lt;groupId&gt;io.github.seccoding&lt;/groupId&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&lt;artifactId&gt;Excel4J&lt;/artifactId&gt;
&nbsp;&nbsp;&nbsp;&nbsp;&lt;version&gt;2.1.1&lt;/version&gt;
	&lt;/dependency&gt;
   </pre>
   
### maven dependency에 Excel4J-2.1.1.jar 파일을 추가할 경우 
<a href="#바로가기">상위로 가기</a>
1. Excel4J-2.1.1.jar파일을 C:\에 복사합니다.
1. Maven 명령어를 이용해 .m2 Repository 에 Excel4J-2.1.1.jar 를 설치(저장)합니다.<pre>mvn install:install-file -Dfile=C:\Excel4J-2.1.1.jar -DgroupId=io.github.seccoding -DartifactId=Excel4J -Dversion=2.1.1 -Dpackaging=jar</pre>
1. 본인의 Project/pom.xml 에 dependency를 추가합니다.<pre>
	&lt;dependency&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;groupId&gt;io.github.seccoding&lt;/groupId&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;artifactId&gt;Excel4J&lt;/artifactId&gt;
	&nbsp;&nbsp;&nbsp;&nbsp;&lt;version&gt;2.1.1&lt;/version&gt;
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
<pre>
import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.annotations.Require;
import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.ExcelRead;

public class ExcelReadTest {

	private static ReadOption ro = new ReadOption();
	private static String filePath = "Excel File Path";
	
	public static void main(String[] args) {

		ro.setFilePath(filePath);
		
		test1();
		test2();
		test3();
		test4();
		test5();
		test6();
	}

	@Deprecated
	public static void test1() {
		System.out.println("test1");
		ro.setSheetName("Sheet1");
		ro.setOutputColumns("B", "C", "D");
		ro.setStartRow(1);
		Map&lt;String, String> result = new ExcelRead().read(ro);

		System.out.println(result);
		System.out.println(result.get("B7"));
		System.out.println(result.get("C7"));
		System.out.println(result.get("D7"));
	}

	@Deprecated
	public static void test2() {
		System.out.println("test2");
		
		ro.setOutputColumns(null); // TestClass의 @Field로 대체
		ro.setSheetName(null); // TestClass의 @ExcelSheet() 로 대체
		ro.setStartRow(0); // TestClass의 @ExcelSheet() 로 대체
		TestClass result = new ExcelRead&lt;TestClass>().readToObject(ro, TestClass.class);

		System.out.println(result.getNo().size());
		System.out.println(result.getColumnName().size());
		System.out.println(result.getType().size());
	}
	
	public static void test3() {
		System.out.println("test3");
		
		ro.setOutputColumns(null); // TestClass의 @Field로 대체
		ro.setSheetName(null); // TestClass의 @ExcelSheet() 로 대체
		ro.setStartRow(0); // TestClass의 @ExcelSheet() 로 대체
		
		List&lt;TestClass2> result = new ExcelRead&lt;TestClass2>().readToList(ro, TestClass2.class);
		System.out.println(result.size());
	}

	@Deprecated
	public static void test4() {
		System.out.println("test4");
		ro.setSheetName("Sheet1");
		String result = new ExcelRead().getValue(ro, "B3");
		System.out.println(result);
	}
	
	public static void test5() {
		System.out.println("test5");
		String result = new ExcelRead().getValue(filePath, "Sheet1", "B3");
		System.out.println(result);
	}
	
	public static void test6() {
		System.out.println("test6");
		String result = new ExcelRead().getValue(filePath, "B3");
		System.out.println(result);
	}

	@ExcelSheet(value="Sheet1", startRow=1)
	@Deprecated
	public static class TestClass {

		@Field("B")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private List&lt;String> no;

		@Field("C")
		private List&lt;String> columnName;

		@Field("D")
		private List&lt;String> type;

		public List&lt;String> getNo() {
			return no;
		}

		public void setNo(List&lt;String> no) {
			this.no = no;
		}

		public List&lt;String> getColumnName() {
			return columnName;
		}

		public void setColumnName(List&lt;String> columnName) {
			this.columnName = columnName;
		}

		public List&lt;String> getType() {
			return type;
		}

		public void setType(List&lt;String> type) {
			this.type = type;
		}

	}

	@ExcelSheet(value="Sheet1", startRow=1)
	public static class TestClass2 {

		@Field("B")
		@Require // 값이 항상 존재하는 컬럼을 지정. 탐색 ROW를 지정할 때 사용.
		private String no;

		@Field("C")
		private String columnName;

		@Field("D")
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
</pre>

---
## Excel File 쓰기 
<a href="#바로가기">상위로 가기</a>
<pre>

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import io.github.seccoding.excel.annotations.Field;
import io.github.seccoding.excel.option.WriteOption;
import io.github.seccoding.excel.write.ExcelWrite;

/**
 * ExcelWriteTest Example
 * 
 * @author Min Chang Jang (mcjang1116@gmail.com)
 */
public class ExcelWriteTest {

	public static void main(String[] args) {

		WriteOption&lt;TestVO> wo = new WriteOption&lt;TestVO>();
		wo.setSheetName("Test");
		wo.setFileName("test.xlsx");
		wo.setFilePath("C:\\Users\\mcjan\\Desktop");

		List&lt;String> titles = new ArrayList&lt;String>();
		titles.add("Title1");
		titles.add("Title2");
		titles.add("Title3");
		titles.add("Title4");
		wo.setTitles(titles);

		List&lt;TestVO> contents = new ArrayList&lt;TestVO>();
		contents.add(new TestVO(1, "ABC", true, "=1+1"));
		contents.add(new TestVO(2, "DEF", true, "=2+2"));
		contents.add(new TestVO(3, "HIJ", true, "=3+3"));
		wo.setContents(contents);

		File excelFile = ExcelWrite.write(wo);
	}

	public static class TestVO {
		
		@Field("Title1")
		private int id;
		
		@Field("Title2")
		private String content;
		
		@Field("Title3")
		private boolean isTrue;
		
		@Field("Title4")
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

</pre>

<a href="#바로가기">상위로 가기</a>