package io.github.seccoding.excel.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class InstanceUtil {

	public static <T> T createInstance(Class<T> resultClass) {
		try {
			Constructor<T> defaultConstructor = resultClass.getDeclaredConstructor();
			defaultConstructor.setAccessible(true);
			return defaultConstructor.newInstance();
		} catch (NoSuchMethodException | SecurityException | InstantiationException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	public static Object invokeMethod(Object obj, Method method, Object ... args) {
		try {
			return method.invoke(obj, args);
		} catch (IllegalAccessException | InvocationTargetException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
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
	
	public static void setValueOfField(Object obj, Field field, Object value) {
		try {
			field.set(obj, value);
		} catch (IllegalArgumentException | IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
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
