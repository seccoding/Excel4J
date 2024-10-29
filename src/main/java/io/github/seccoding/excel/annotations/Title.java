package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 엑셀 워크시트에 작성할 타이틀 지정.
 * 멤버변수에 지정하며
 * 이 애노테이션이 적용된 멤버변수만 워크시트에 작성한다.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Title {

	/**
	 * 타이틀 이름
	 */
	public String value();
	
	/**
	 * 타이틀을 무시하고 넘어갈지 여부
	 * @Merge 애노테이션이 적용된 경우 사용한다.
	 */
	public boolean ignoreTitle() default false;
	/**
	 * 이전 줄에 타이틀을 쓸지 여부
	 * @Merge 애노테이션이 적용된 경우 사용한다.
	 */
	public boolean appendPrevRow() default false;
	
	public Merge merge() default @Merge;
}
