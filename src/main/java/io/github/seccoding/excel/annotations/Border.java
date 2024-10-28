package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Border {

	public BorderStyle value() default BorderStyle.NONE;
	
	public IndexedColors color() default IndexedColors.BLACK;
	
	public boolean top() default true;
	public boolean right() default true;
	public boolean bottom() default true;
	public boolean left() default true;
	
}
