package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.IndexedColors;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface BackgroundColor {

	public IndexedColors value();
	
}
