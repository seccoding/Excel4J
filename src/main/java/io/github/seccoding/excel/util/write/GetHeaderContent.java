package io.github.seccoding.excel.util.write;

import java.lang.reflect.Field;
import java.util.List;

import io.github.seccoding.excel.util.write.share.WriteShare;

public class GetHeaderContent {

	public static String getParentTitle(Object obj, Field[] fields, String fieldName) {
		
		for (Field field : fields) {
			field.setAccessible(true);
			if ( field.getName().equals(fieldName) ) {
				try {
					return (String) field.get(obj);
				} catch (IllegalArgumentException | IllegalAccessException e) {
					throw new RuntimeException(e.getMessage(), e);
				}
			}
		}
		
		return "";
	}
	
	public static Object getFirstContent() {
		List<?> values = WriteShare.writeOption.getContents();
		
		if ( values != null && values.size() > 0 ) {
			return values.get(0);
		}
		
		return null;
	}
	
}
