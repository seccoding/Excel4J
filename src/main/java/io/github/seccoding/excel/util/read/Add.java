package io.github.seccoding.excel.util.read;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class Add {

	private static final int LIST = 1;
	private static final int MAP = 2;
	private static final int SET = 3;
	private static final int STRING = 4;
	private static final int BYTE = 5;
	private static final int SHORT = 6;
	private static final int INTEGER = 7;
	private static final int LONG = 8;
	private static final int FLOAT = 9;
	private static final int DOUBLE = 10;
	private static final int BOOLEAN = 10;
	
	private static final int NONE = -1;
	
	private static final Map<Class<?>, Integer> SUPPORT_TYPES;
	
	static {
		SUPPORT_TYPES = new HashMap<Class<?>, Integer>();
		SUPPORT_TYPES.put(List.class, LIST);
		SUPPORT_TYPES.put(Map.class, MAP);
		SUPPORT_TYPES.put(Set.class, SET);
		SUPPORT_TYPES.put(String.class, STRING);
		SUPPORT_TYPES.put(byte.class, BYTE);
		SUPPORT_TYPES.put(short.class, SHORT);
		SUPPORT_TYPES.put(int.class, INTEGER);
		SUPPORT_TYPES.put(long.class, LONG);
		SUPPORT_TYPES.put(float.class, FLOAT);
		SUPPORT_TYPES.put(double.class, DOUBLE);
		SUPPORT_TYPES.put(boolean.class, BOOLEAN);
	}
	
	
	@SuppressWarnings({"rawtypes", "unchecked"})
	public static void add(String fieldName, Object obj, String key, String value) {
		
		Map<String, Object> fieldAndCollection = getFieldAndObject(fieldName, obj);
		Object collection = fieldAndCollection.get("COLLECTION");
		Field f = (Field) fieldAndCollection.get("FIELD");
		
		if ( collection instanceof List ) {
			List list = (List) collection;
			set(f, obj, list);
			list.add(value);
			return;
		}
		if ( collection instanceof Set ) {
			Set set = (Set) collection;
			set(f, obj, set);
			set.add(value);
			return;
		}
		if ( collection instanceof Map ) {
			Map map = (Map) collection;
			set(f, obj, map);
			map.put(key, value);
			return;
		}
		
		if ( SUPPORT_TYPES.containsKey(collection.getClass()) ) {
			set(f, obj, value);
			return;
		}
		
	}
	
	private static Map<String, Object> getFieldAndObject(String fieldName, Object obj) {
		
		try {
			Field f = getField(fieldName, obj);
			Object collection = f.get(obj);
			collection = getOrMakeCollection(f, collection);
			
			Map<String, Object> result = new HashMap<String, Object>();
			result.put("FIELD", f);
			result.put("COLLECTION", collection);
			
			return result;
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	private static Field getField(String fieldName, Object obj) {
		Field f = null;
		try {
			f = obj.getClass().getDeclaredField(fieldName);
		} catch (NoSuchFieldException e1) {
			throw new RuntimeException(e1.getMessage(), e1);
		} catch (SecurityException e1) {
			throw new RuntimeException(e1.getMessage(), e1);
		}
		
		f.setAccessible(true);
		
		return f;
	}
	
	@SuppressWarnings("rawtypes")
	private static Object getOrMakeCollection(Field f, Object collection) {
		
		int collectionType = getCollectionType(f);
		
		if ( collectionType == LIST ) {
			List list = (List) collection;
			if ( list == null ) {
				list = new ArrayList();
			}
			return list;
		}
		if ( collectionType == SET ) {
			Set set = (Set) collection;
			if ( set == null ) {
				set = new HashSet();
			}
			return set;
		}
		if ( collectionType == MAP ) {
			Map map = (Map) collection;
			if ( map == null ) {
				map = new HashMap();
			}
			return map;
		}
		if ( collectionType == STRING ) {
			String str = String.valueOf(collection);
			if ( str == null ) {
				str = "";
			}
			return str;
		}
		if ( collectionType == BYTE ) {
			Byte bt = Byte.parseByte( String.valueOf(collection) );
			return bt;
		}
		if ( collectionType == SHORT ) {
			Short sht = Short.parseShort(String.valueOf(collection));
			return sht;
		}
		if ( collectionType == INTEGER ) {
			Integer intg = Integer.parseInt(String.valueOf(collection));
			return intg;
		}
		if ( collectionType == LONG ) {
			Long lng = Long.parseLong(String.valueOf(collection));
			return lng;
		}
		if ( collectionType == FLOAT ) {
			Float flt = Float.parseFloat(String.valueOf(collection));
			return flt;
		}
		if ( collectionType == DOUBLE ) {
			Double dbl = Double.parseDouble(String.valueOf(collection));
			return dbl;
		}
		if ( collectionType == BOOLEAN ) {
			Boolean bool = Boolean.parseBoolean(String.valueOf(collection));
			return bool;
		}
		
		throw new RuntimeException(f.getType() + " in not support.");
		
	}
	
	private static int getCollectionType(Field f) {
		Class<?> type = f.getType();
		return SUPPORT_TYPES.containsKey(type) ? SUPPORT_TYPES.get(type) : NONE;
	}
	
	private static void set(Field f, Object obj, Object collection) {
		try {
			f.set(obj, collection);
		} catch (IllegalArgumentException e) {
			throw new RuntimeException(e.getMessage(), e);
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
}
