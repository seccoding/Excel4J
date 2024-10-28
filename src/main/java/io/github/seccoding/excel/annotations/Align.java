package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Align {
	
	public HorizontalAlignment value() default HorizontalAlignment.LEFT;
	public VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;
	
}
