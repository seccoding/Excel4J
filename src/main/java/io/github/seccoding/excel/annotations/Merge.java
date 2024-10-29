package io.github.seccoding.excel.annotations;

/**
 * Title에서 사용하는 서브 애노테이션.
 * 엑셀 워크 시트 작성할 때 타이틀 컬럼의 병합을 담당.
 */
public @interface Merge {

	/**
	 * 병합된 셀 내용 지정.
	 * 공백일 경우 @Title의 value를 그대로 사용. 
	 */
	public String value() default "";
	
	/**
	 * 줄 병합할 개수.
	 */
	public int rows() default 0;
	/**
	 * 칸 병합할 개수.
	 */
	public int cols() default 0;
	
}
