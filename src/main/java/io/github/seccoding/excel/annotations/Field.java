package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 엑셀 워크 시트를 읽을 때 사용.
 * 컬럼의 이름만 입력
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Field {

	/**
	 * 읽을 컬럼의 이름
	 */
	public String value();
	
	/**
	 * 컬럼이 날짜 타입일 때 true 지정.
	 * 날짜 타입에 대해 false로 지정할 경우 숫자로 반환된다.
	 */
	public boolean isDate() default false;
	
	/**
	 * 날짜타입의 컬럼을 읽을 때 사용할 날짜 포멧
	 * 기본값 : 연-월-일
	 */
	public String dateFormat() default "yyyy-MM-dd";
	
}
