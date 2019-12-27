package com.ss.poi.util;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonPrimitive;
import com.google.gson.JsonSerializationContext;
import com.google.gson.JsonSerializer;
import com.google.gson.reflect.TypeToken;

public class GsonUtil {

	/** 文件转字符串
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static String readFile2String(String filePath) throws IOException {
		long startTime = new Date().getTime();
		FileInputStream inStream = new FileInputStream(filePath);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		byte[] buffer = new byte[1024];
		int length = -1;
		while ((length = inStream.read(buffer)) != -1) {
			bos.write(buffer, 0, length);
		}
		bos.close();
		inStream.close();
		//System.out.println(bos.toString());
		System.out.println("读取文件完成！耗时："+(new Date().getTime()-startTime)+"ms");
		return bos.toString();
	}
	
	/** 从传入的json中通过父类key查找value
	 * @param json
	 * @param key
	 * @return
	 */
	public static Object getValForJson(String json,String key) {
		long startTime = new Date().getTime();
		//LinkedHashMap<String, String> jsonMap = JSON.parseObject(json, new TypeReference<LinkedHashMap<String, String>>(){});
        Map<?, ?> map = new Gson().fromJson(json, Map.class);
		//System.out.println("字符串转Map："+jsonMap);
		System.out.println("解析json完成！耗时："+(new Date().getTime()-startTime)+"ms");
		return map.get(key);
	}
	/**
	 * json to javabean
	 * 
	 * @param json
	 */
	public static <T> T jsonToObject(String json, Class<T> classOfT) {
		try {
			Gson gson = new Gson();
			return gson.fromJson(json, classOfT);
		} catch (Exception e) {
			return null;
		}
	}

	public static <T> List<T> jsonToList(String json, Class<T[]> clazz) {

		Gson gson = new Gson();
		T[] arr = gson.fromJson(json, clazz);
		return Arrays.asList(arr);
	}

	public static <T> Map<String, T> jsonToSMap(String json, Class<T> classOfT) {
		Gson gson = new Gson();
		Map<String, T> maps = gson.fromJson(json,
				new TypeToken<Map<String, T>>() {
				}.getType());
		return maps;
	}

	/**
	 * javabean to json
	 * 
	 * @param
	 * @return
	 */
	public static String objectToJson(Object obj) {
		if (obj instanceof String) {
			return (String) obj;
		} else {
			Gson gson = new GsonBuilder().disableHtmlEscaping().create();
			String json = gson.toJson(obj);
			return json;
		}
	}

	/**
	 * list to json
	 * 
	 * @param list
	 * @return
	 */
	public static String listToJson(List<?> list) {

		Gson gson = new GsonBuilder().registerTypeAdapter(Double.class,
				new JsonSerializer<Double>() {
					/*@Override*/
					public JsonElement serialize(Double src, Type typeOfSrc,
							JsonSerializationContext context) {
						if (src == src.longValue())
							return new JsonPrimitive(src.longValue());
						return new JsonPrimitive(src);
					}
				}).create();
		String json = gson.toJson(list);
		return json;
	}

	/**
	 * map to json
	 * 
	 * @param map
	 * @return
	 */
	public static String mapToJson(Map<String, Object> map) {

		Gson gson = new Gson();
		String json = gson.toJson(map);
		return json;
	}

}
