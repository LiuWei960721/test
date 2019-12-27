package com.ss.poi.test;

import com.ss.poi.controller.WordReport;
import com.ss.poi.util.GsonUtil;

/**
 * @Description 
 */
public class Jsontest {
	public static void main(String[] args) throws Exception {

		String dataJson = GsonUtil.readFile2String("G:/workspace/java/elecmoni3.0/target/TestResult/123/0/0_高压初测流程_低压调试框-灯丝正常电流测试（负载）_低压开灯丝电流正常负载测试示波器读取/20190516-102456.json");
		
		WordReport.creatWordReport(dataJson);

	}

}
