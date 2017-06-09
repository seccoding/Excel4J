package io.github.seccoding.excel;

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
		
		List<String[]> contents = new ArrayList<String[]>();
		contents.add(new String[]{"1", "2", "3"});
		contents.add(new String[]{"11", "22", "33"});
		contents.add(new String[]{"111", "222", "333"});
		wo.setContents(contents);
		
		File excelFile = ExcelWrite.write(wo);
	}
	
}
