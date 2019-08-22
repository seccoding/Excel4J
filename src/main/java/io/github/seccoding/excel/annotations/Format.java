package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Format {

	public String alignment() default LEFT;
	public String verticalAlignment() default V_CENTER;
	public boolean bold() default false;
	
	public String dataFormat() default "";
	
	/**
	 * <pre>
	 * 변경하고자 하는 포멧
	 * 
	 * &#64;Field.date 의 값이 true일 때만 사용.
	 * </pre>
	 */
	public String toDataFormat() default "";
	
	public static final String RIGHT = "RIGHT";
	public static final String LEFT = "LEFT";
	public static final String CENTER = "CENTER";
	
	public static final String V_TOP = "V_TOP";
	public static final String V_BOTTOM = "V_BOTTOM";
	public static final String V_CENTER = "V_CENTER";
	
}
