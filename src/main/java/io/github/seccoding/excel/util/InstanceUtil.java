package io.github.seccoding.excel.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * 인스턴스 리플렉션 유틸리티
 */
public class InstanceUtil {

	/**
	 * 새로운 인스턴스를 만든다.
	 * @param <T> 인스턴스 타입
	 * @param resultClass 인스턴스 클래스 원본
	 * @return 새로운 인스턴스
	 */
	public static <T> T createInstance(Class<T> resultClass) {
		try {
			Constructor<T> defaultConstructor = resultClass.getDeclaredConstructor();
			defaultConstructor.setAccessible(true);
			return defaultConstructor.newInstance();
		} catch (NoSuchMethodException | SecurityException | InstantiationException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	/**
	 * 메소드를 실행한다.
	 * @param obj 메소드를 실행할 인스턴스
	 * @param method 실행할 메소드
	 * @param args 메소드 파라미터
	 * @return 메소드 실행 결과
	 */
	public static Object invokeMethod(Object obj, Method method, Object ... args) {
		try {
			return method.invoke(obj, args);
		} catch (IllegalAccessException | InvocationTargetException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	/**
	 * 인스턴스에서 메소드를 추출한다.
	 * @param obj 메소드를 추출할 인스턴스
	 * @param prefix 메소드 시작 이름
	 * @param fieldName prefix를 제외한 메소드 이름
	 * @param parameterTypes 메소드 파라미터 타입
	 * @return 메소드
	 */
	public static Method getMethod(Object obj, String prefix, String fieldName, Class<?> ... parameterTypes) {
		String methodName = prefix + (fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1));
		try {
			Method method = obj.getClass().getDeclaredMethod(methodName, parameterTypes);
			method.setAccessible(true);
			return method;
		} catch (NoSuchMethodException | SecurityException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	/**
	 * 인스턴스에서 멤버변수를 추출한다.
	 * @param obj 멤버변수를 추출할 인스턴스
	 * @param fieldName 추출할 멤버변수 이름
	 * @return 멤버변수 필드
	 */
	public static Field getFieldInNewInstance(Object obj, String fieldName) {
		try {
			Field field = obj.getClass().getDeclaredField(fieldName);
			field.setAccessible(true);
			return field;
		} catch (NoSuchFieldException | SecurityException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
}
