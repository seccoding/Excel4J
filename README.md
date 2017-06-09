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