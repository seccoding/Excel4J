package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Title {

	public String value();

	public String parentTitle() default "";
	public int parentRowMerge() default 1;
	public int parentCellMerge() default 1;
	
	public int rowMerge() default 1;
	public int cellMerge() default 1;
	
	public int sort() default 0;
	
	public boolean date() default false;
	
}
