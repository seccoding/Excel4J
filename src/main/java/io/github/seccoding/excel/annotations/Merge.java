package io.github.seccoding.excel.annotations;

public @interface Merge {

	public String value() default "";
	
	public int rows() default 0;
	public int cols() default 0;
	
}
