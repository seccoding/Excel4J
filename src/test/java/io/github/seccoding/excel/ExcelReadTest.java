package io.github.seccoding.excel;

import java.util.List;
import java.util.Map;

import io.github.seccoding.excel.option.ReadOption;
import io.github.seccoding.excel.read.ExcelRead;

public class ExcelReadTest {

	public static void main(String[] args) {

		ReadOption ro = new ReadOption();
		ro.setFilePath("/Users/mcjang/ktest/uploadedFile/practiceTest.xlsx");
		ro.setOutputColumns("C", "D", "E", "F", "G", "H", "I");
		ro.setStartRow(3);

		List<Map<String, String>> result = ExcelRead.read(ro);

		for (Map<String, String> map : result) {
			System.out.println(map.get("E"));
		}
	}

}
