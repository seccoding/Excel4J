package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 엑셀 작성시 셀 정렬 담당.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Align {
	
	/**
	 * 수평 정렬
	 * 기본값: 왼쪽 정렬
	 */
	public HorizontalAlignment value() default HorizontalAlignment.LEFT;
	
	/**
	 * 수직 정렬
	 * 기본값: 중앙 정렬
	 */
	public VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;
	
}
