# Excel
Java 에서 엑셀파일을 읽고 쓰는 유틸리티<br/>
xls 와 xlsx를 모두 지원함.


## 사용 방법
### maven dependency에 Excel-1.0.0.jar 파일을 추가할 경우
1. Excel-1.0.0.jar파일을 C:\에 복사합니다.
1. Maven 명령어를 이용해 .m2 Repository 에 Excel-1.0.0.jar 를 설치(저장)합니다.
1. <pre>
mvn install:install-file -Dfile=C:\Excel-1.0.0.jar -DgroupId=io.github.seccoding -DartifactId=Excel -Dversion=1.0.0 -Dpackaging=jar
</pre>
1. 본인의 Project/pom.xml 에 dependency를 추가합니다.
1. <pre>
	&lt;dependency&gt;
		&lt;groupId&gt;io.github.seccoding&lt;/groupId&gt;
		&lt;artifactId&gt;Excel&lt;/artifactId&gt;
		&lt;version&gt;1.0.0&lt;/version&gt;
	&lt;/dependency&gt;
</pre>

### 소스코드를 사용할 경우
1. Clone or Download 를 클릭합니다.
1. Download ZIP 을 클릭해 소스코드를 다운로드 받습니다.
1. Excel/pom.xml의 dependencies를 본인의 Project/pom.xml 에 붙혀넣습니다.
1. Excel/src 이하의 자바코드를 본인의 Project에 붙혀넣습니다. 
---
## Excel File 읽기
<pre>
import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.ExcelRead;

public class ExcelReadTest {

	public static void main(String[] args) {

		ReadOption ro = new ReadOption();
		ro.setFilePath("/Users/mcjang/ktest/uploadedFile/excelFile.xlsx");
		ro.setOutputColumns("C", "D", "E", "F", "G", "H", "I");
		ro.setStartRow(3);

		List&lt;Map&lt;String, String&gt;&gt; result = ExcelRead.read(ro);

		for (Map&lt;String, String&gt; map : result) {
			System.out.println(map.get("E"));
		}
	}

}
</pre>

---
## Excel File 쓰기
<pre>
import java.io.File;
import java.util.ArrayList;
import java.util.List;

import io.github.seccoding.excel.option.WriteOption;
import io.github.seccoding.excel.write.ExcelWrite;

public class ExcelWriteTest {

	public static void main(String[] args) {
		WriteOption wo = new WriteOption();
		wo.setSheetName("Test");
		wo.setFileName("test.xlsx");
		wo.setFilePath("d:\\");
		
		List<String> titles = new ArrayList<String>();
		titles.add("Title1");
		titles.add("Title2");
		titles.add("Title3");
		wo.setTitles(titles);
		
		List&lt;String[]&gt; contents = new ArrayList&lt;String[]&gt;();
		contents.add(new String[]{"1", "2", "3"});
		contents.add(new String[]{"11", "22", "33"});
		contents.add(new String[]{"111", "222", "333"});
		wo.setContents(contents);
		
		File excelFile = ExcelWrite.write(wo);
	}
	
}
</pre>