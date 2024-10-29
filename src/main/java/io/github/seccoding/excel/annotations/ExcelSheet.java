package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 엑셀 워크시트 지정.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {

	/**
	 * 엑셀 워크시트의 이름
	 * @return
	 */
	public String value() default "";
	
	/**
	 * 읽거나 쓰기 시작하는 줄 번호 (0-base)
	 */
	public int startRow() default 1;
	
}
