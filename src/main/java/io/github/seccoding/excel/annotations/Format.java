package io.github.seccoding.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import org.apache.poi.ss.usermodel.CellStyle;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Format {

	public short alignment() default LEFT;
	public short verticalAlignment() default V_CENTER;
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
	
	public static final short RIGHT = CellStyle.ALIGN_RIGHT;
	public static final short LEFT = CellStyle.ALIGN_LEFT;
	public static final short CENTER = CellStyle.ALIGN_CENTER;
	
	public static final short V_TOP = CellStyle.VERTICAL_TOP;
	public static final short V_BOTTOM = CellStyle.VERTICAL_BOTTOM;
	public static final short V_CENTER = CellStyle.VERTICAL_CENTER;
	
}
