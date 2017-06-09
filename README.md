# Excel
Java 에서 엑셀파일을 읽고 쓰는 유틸리티<br/>
xls 와 xlsx를 모두 지원함.

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