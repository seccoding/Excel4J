package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 엑셀 워크 시트의 폰트 스타일 지정 담당.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Text {

	/**
	 * 셀 텍스트를 굵게 표현할지 여부.
	 */
	public boolean bold() default false;
	/**
	 * 셀 텍스트의 색상 지정.
	 * 기본값 : 검정.
	 */
	public IndexedColors color() default IndexedColors.BLACK;
	
}
