package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 엑셀 워크시트 내 모든 셀에 대한 경계선 지정.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Border {

	/**
	 * 경계선 종류
	 * 기본값: 없음.
	 */
	public BorderStyle value() default BorderStyle.NONE;
	
	/**
	 * 경계선 색상.
	 * 기본값: 검정
	 */
	public IndexedColors color() default IndexedColors.BLACK;
	
	/**
	 * 셀 위쪽 경계선 지정 여부
	 * 기본값: 경계선 지정함.
	 */
	public boolean top() default true;
	/**
	 * 셀 오른쪽 경계선 지정 여부
	 * 기본값: 경계선 지정함.
	 */
	public boolean right() default true;
	/**
	 * 셀 아래쪽 경계선 지정 여부
	 * 기본값: 경계선 지정함.
	 */
	public boolean bottom() default true;
	/**
	 * 셀 왼쪽 경계선 지정 여부
	 * 기본값: 경계선 지정함.
	 */
	public boolean left() default true;
	
}
