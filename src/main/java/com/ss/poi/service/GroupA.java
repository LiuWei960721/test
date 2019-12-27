package com.ss.poi.service;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.ss.poi.entity.EPCBusVoltEnum;
import com.ss.poi.entity.EPCVoltStageEnum;
import com.ss.poi.util.ImgUtil;
import com.ss.poi.util.OperateTable;
import com.ss.poi.util.ReportFileUtil;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import com.ss.poi.entity.TestResultData;
import com.ss.poi.util.POITableUtil;

import java.util.Map;
import java.util.Set;

/**
 * @Description 
 */
public class GroupA {

	//母线电流100030
	public static void genFunc100030(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, data.getEPCStage().getVoltStageStr()+data.getEPCStage().getBusVoltStr() +"电流值");
		POITableUtil.setCellValueString(table, 0, 2, "母线功耗");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值");
		//接口单
		switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getVoltStageStr())){
			case 低压:
				POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("lowBusVolt").getAsDouble());
				break;
			case 动态:
				POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("dynamicBusVolt").getAsDouble());
				break;
			case 静态:
				POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("stateBusVolt").getAsDouble());
				break;
				case 不需要设置:
			default:
				break;
		}

		switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getVoltStageStr())){
			case 低压:
				POITableUtil.setCellValueDouble(table, 1, 2, data.getTestInputData().get("lowBusPow").getAsDouble());
				break;
			case 动态:
				POITableUtil.setCellValueDouble(table, 1, 2, data.getTestInputData().get("dynamicBusPow").getAsDouble());
				break;
			case 静态:
				POITableUtil.setCellValueDouble(table, 1, 2, data.getTestInputData().get("stateBusPow").getAsDouble());
				break;
			case 不需要设置:
			default:
				break;
		}

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("Curr").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("Power").getAsDouble());
		doc.createParagraph();
	}

	//遥测电压100040
	public static void genFunc100040(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 6);
		POITableUtil.setTableColmnWidth(table, new int[]{1200,1200,1200,1200,1200,1200});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "螺流遥测");
		POITableUtil.setCellValueString(table, 0, 2, "阳压遥测");
		POITableUtil.setCellValueString(table, 0, 3, "开关机状态遥测");
		POITableUtil.setCellValueString(table, 0, 4, "自动重启遥测");
		POITableUtil.setCellValueString(table, 0, 5, "母线电流遥测");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");//yanYaYaoCeVoltMin
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("luoLiuYaoCeVoltMin").getAsString()+"~"+data.getTestInputData().get("luoLiuYaoCeVoltMax").getAsString());
		POITableUtil.setCellValueString(table, 1, 2, data.getTestInputData().get("yanYaYaoCeVoltMin").getAsString()+"~"+data.getTestInputData().get("yanYaYaoCeVoltMax").getAsString());
		POITableUtil.setCellValueString(table, 1, 3, data.getTestInputData().get("kaiGuanJiYaoCeVoltMin").getAsString()+"~"+data.getTestInputData().get("kaiGuanJiYaoCeVoltMax").getAsString());
		POITableUtil.setCellValueString(table, 1, 4, data.getTestInputData().get("ziDongChongQiYaoCeVoltMin").getAsString()+"~"+data.getTestInputData().get("ziDongChongQiYaoCeVoltMax").getAsString());
		POITableUtil.setCellValueString(table, 1, 5, data.getTestInputData().get("muXianYaoCeVoltMin").getAsString()+"~"+data.getTestInputData().get("muXianYaoCeVoltMin").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("LuoLIuYaoCe").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("YangYaYaoCe").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 3, data.getTestExportData().get("KaiGuanjiYaoCe").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 4, data.getTestExportData().get("ZiDongChongQiYaoCe").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 5, data.getTestExportData().get("MuXianDianLiuYaoCe").getAsDouble());
		doc.createParagraph();

	}

	//BoostBuck电压测试100050
	public static void genFunc100050(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "boost-buck电压测试值");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("BoostBuckMin").getAsString()+"~"+data.getTestInputData().get("BoostBuckMax").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("BoostBuckYaoCe").getAsDouble());
		doc.createParagraph();
	}

	//CAMP电压电流测试 = 100060,
	public static void genFunc100060(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 5, 4);
		POITableUtil.setTableColmnWidth(table, new int[]{3000,1500,1500,1500});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "Camp+5");
		POITableUtil.setCellValueString(table, 0, 2, "Camp+12");
		POITableUtil.setCellValueString(table, 0, 3, "Camp-12");

		POITableUtil.setCellValueString(table, 1, 0, "指标电压要求值（V）");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("zhengV5V").getAsString());
		POITableUtil.setCellValueString(table, 1, 2, data.getTestInputData().get("zhengV12V").getAsString());
		POITableUtil.setCellValueString(table, 1, 3, data.getTestInputData().get("fuV12V").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "指标电流要求值（A）");
		POITableUtil.setCellValueString(table, 2, 1, data.getTestInputData().get("zhengI5V").getAsString());
		POITableUtil.setCellValueString(table, 2, 2, data.getTestInputData().get("zhengI12V").getAsString());
		POITableUtil.setCellValueString(table, 2, 3, data.getTestInputData().get("fuI12V").getAsString());

		POITableUtil.setCellValueString(table, 3, 0, "实际电压测试值（V）");
		POITableUtil.setCellValueDouble(table, 3, 1, data.getTestExportData().get("VZheng5").getAsDouble());
		POITableUtil.setCellValueDouble(table, 3, 2, data.getTestExportData().get("VZheng12").getAsDouble());
		POITableUtil.setCellValueDouble(table, 3, 3, data.getTestExportData().get("VFu12").getAsDouble());

		POITableUtil.setCellValueString(table, 4, 0, "实际电流测试值（A）");
		POITableUtil.setCellValueDouble(table, 4, 1, data.getTestExportData().get("IZheng5").getAsDouble());
		POITableUtil.setCellValueDouble(table, 4, 2, data.getTestExportData().get("IZheng12").getAsDouble());
		POITableUtil.setCellValueDouble(table, 4, 3, data.getTestExportData().get("IFu12").getAsDouble());
		doc.createParagraph();
	}

	//deltH电压测试 = 100070,
	public static void genFunc100070(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "deltH电压测试值");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("deltHMin").getAsString()+"~"+data.getTestInputData().get("deltHMax").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("DeltHVolt").getAsDouble());
		doc.createParagraph();
	}

	//HC1C2   100080
	public static void genFunc100080(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 8, 5);
		POITableUtil.setTableColmnWidth(table, new int[]{1000,1000,1000,1000,1000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
		POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
		POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
		POITableUtil.setCellValueString(table, 0, 4, "与标称值相差值（V）");

		POITableUtil.setCellValueString(table, 1, 0, "螺旋级（H）");
		POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("setVoltH").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("testVoltH").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 3, data.getTestExportData().get("actVoltH").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 4, data.getTestExportData().get("actVoltHDelt").getAsDouble());

		POITableUtil.setCellValueString(table, 2, 0, "阳极（A）");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("setVoltA1").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("testVoltA1").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 3, data.getTestExportData().get("actVoltA1").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 4, data.getTestExportData().get("actVoltA1Delt").getAsDouble());

		POITableUtil.setCellValueString(table, 3, 0, "收集极（C1）");
		POITableUtil.setCellValueDouble(table, 3, 1, data.getTestExportData().get("setVoltC1").getAsDouble());
		POITableUtil.setCellValueDouble(table, 3, 2, data.getTestExportData().get("testVoltC1").getAsDouble());
		POITableUtil.setCellValueDouble(table, 3, 3, data.getTestExportData().get("actVoltC1").getAsDouble());
		POITableUtil.setCellValueDouble(table, 3, 4, data.getTestExportData().get("actVoltC1Delt").getAsDouble());

		POITableUtil.setCellValueString(table, 4, 0, "收集极（C2）");
		POITableUtil.setCellValueDouble(table, 4, 1, data.getTestExportData().get("setVoltC2").getAsDouble());
		POITableUtil.setCellValueDouble(table, 4, 2, data.getTestExportData().get("testVoltC2").getAsDouble());
		POITableUtil.setCellValueDouble(table, 4, 3, data.getTestExportData().get("actVoltC2").getAsDouble());
		POITableUtil.setCellValueDouble(table, 4, 4, data.getTestExportData().get("actVoltC2Delt").getAsDouble());

		POITableUtil.setCellValueString(table, 5, 0, "收集极（C3）");
		POITableUtil.setCellValueDouble(table, 5, 1, data.getTestExportData().get("setVoltC3").getAsDouble());
		POITableUtil.setCellValueDouble(table, 5, 2, data.getTestExportData().get("testVoltC3").getAsDouble());
		POITableUtil.setCellValueDouble(table, 5, 3, data.getTestExportData().get("actVoltC3").getAsDouble());
		POITableUtil.setCellValueDouble(table, 5, 4, data.getTestExportData().get("actVoltC3Delt").getAsDouble());

		POITableUtil.setCellValueString(table, 6, 0, "收集极（C4）");
		POITableUtil.setCellValueDouble(table, 6, 1, data.getTestExportData().get("setVoltC4").getAsDouble());
		POITableUtil.setCellValueDouble(table, 6, 2, data.getTestExportData().get("testVoltC4").getAsDouble());
		POITableUtil.setCellValueDouble(table, 6, 3, data.getTestExportData().get("actVoltC4").getAsDouble());
		POITableUtil.setCellValueDouble(table, 6, 4, data.getTestExportData().get("actVoltC4Delt").getAsDouble());

		POITableUtil.setCellValueString(table, 7, 0, "阴极（K）");
		POITableUtil.setCellValueDouble(table, 7, 1, data.getTestExportData().get("setVoltK").getAsDouble());
		POITableUtil.setCellValueDouble(table, 7, 2, data.getTestExportData().get("testVoltK").getAsDouble());
		POITableUtil.setCellValueDouble(table, 7, 3, data.getTestExportData().get("actVoltK").getAsDouble());
		POITableUtil.setCellValueDouble(table, 7, 4, data.getTestExportData().get("actVoltKDelt").getAsDouble());
		doc.createParagraph();
	}

	//电阻电容检查与设置 = 100130
	public static void genFunc100130(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 5, 4);
		POITableUtil.setTableColmnWidth(table, new int[]{1000,1000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "序号");
		POITableUtil.setCellValueString(table, 0, 1, "");
		POITableUtil.setCellValueString(table, 0, 2,"");
		POITableUtil.setCellValueString(table, 0, 3, "程序回读值");

		POITableUtil.setCellValueString(table, 1, 0, "1");
		POITableUtil.setCellValueString(table, 1, 1, "PLC初始安装");
		POITableUtil.setCellValueString(table, 1, 2,"");
		POITableUtil.setCellValueString(table, 1, 3, "");

		POITableUtil.setCellValueString(table, 2, 0, "");
		POITableUtil.setCellValueString(table, 2, 1, "");
		POITableUtil.setCellValueString(table, 2, 2,"boost-buck电阻值");
		POITableUtil.setCellValueDouble(table, 2, 3, data.getTestExportData().get("BoostBuckR").getAsDouble());

		POITableUtil.setCellValueString(table, 3, 0, "");
		POITableUtil.setCellValueString(table, 3, 1, "");
		POITableUtil.setCellValueString(table, 3, 2,"K电阻");
		POITableUtil.setCellValueDouble(table, 3, 3, data.getTestExportData().get("YinjiKR").getAsDouble());

		/*POITableUtil.setCellValueString(table, 4, 0, "");
		POITableUtil.setCellValueString(table, 4, 1, "");
		POITableUtil.setCellValueString(table, 4, 2,"R6R7电阻值");
		POITableUtil.setCellValueDouble(table, 4, 3, data.getTestExportData().get("BoostBuckR").getAsDouble());
*/
		POITableUtil.setCellValueString(table, 4, 0, "");
		POITableUtil.setCellValueString(table, 4, 1, "");
		POITableUtil.setCellValueString(table, 4, 2,"电容值");
		POITableUtil.setCellValueDouble(table, 4, 3, data.getTestExportData().get("XieZhengDianRongR").getAsDouble());

		doc.createParagraph();
	}

	/*//低压开浪涌 = 100160
	public static void genFunc100160(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "低压开浪涌");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
		POITableUtil.setCellValueString(table, 2, 1,  data.getTestExportData().get("VoltProtect").getAsString());
		doc.createParagraph();
	}*/

	//回读电容 = 100180
	public static void genFunc100180(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 2, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "谐振电容值");

		POITableUtil.setCellValueString(table, 1, 0, "程序回读值");
		POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("Capacitance").getAsDouble());

		doc.createParagraph();
	}

	// 高压开浪涌   100190
	public static void genFunc100190(XWPFDocument doc,TestResultData data) throws Exception{
		String base64 = data.getTestExportData().get("imageBase64").getAsString();
		XWPFParagraph para = doc.createParagraph();
		ImgUtil.insertForPara(para,base64,400,400);
		doc.createParagraph();

		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");

		//接口单
		switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getVoltStageStr())){
			case 低压:
				POITableUtil.setCellValueString(table, 0, 1, "低压开浪涌");
				break;
			case 动态:
			case 静态:
				POITableUtil.setCellValueString(table, 0, 1, "高压开浪涌");
				break;
		}
		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("turnonSurge").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
		POITableUtil.setCellValueString(table, 2, 1,  data.getTestExportData().get("turnonSurge").getAsString());


		doc.createParagraph();



	}




	//环路稳定性测试 = 100200  低压
	public static void genFunc100200(XWPFDocument doc,TestResultData data){
		switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getEPCLoadStage()) ){
			case 静态:
			case 动态:
				genFunc100200B(doc,data);
				break;
			case 低压:
			default:
				genFunc100200B(doc,data);
				return;
		}
	}

	//环路稳定性测试 = 100200  低压
	public static void genFunc100200A(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "低压欠压保护值");
		POITableUtil.setCellValueString(table, 0, 2, "低压过压保护值");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueString(table, 1, 2, "接口单");

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
		POITableUtil.setCellValueString(table, 2, 1,  data.getTestExportData().get("VoltProtect").getAsString());
		POITableUtil.setCellValueString(table, 2, 2,  data.getTestExportData().get("VoltProtect").getAsString());
		doc.createParagraph();
	}

	//环路稳定性测试 = 100200   高压
	public static void genFunc100200B(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 2, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "高压环路稳定性测试");

		POITableUtil.setCellValueString(table, 1, 0, "稳定性判断");
		POITableUtil.setCellValueString(table, 1, 1, "环路正常");

		doc.createParagraph();
	}


	//开机阈值测试 = 100230
	public static void genFunc100230(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "低压开机阈值");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("busVoltProtMax").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
		POITableUtil.setCellValueString(table, 2, 1, data.getTestExportData().get("VoltProtect").getAsString());

		doc.createParagraph();
	}


	//谐振波形
	public static void genFunc100240(XWPFDocument doc,TestResultData data) throws Exception {
		String base64 = data.getTestExportData().get("imageBase64").getAsString();
		XWPFParagraph para = doc.createParagraph();
		ImgUtil.insertForPara(para, base64, 400, 400);
		doc.createParagraph();

	}

	//低压开灯丝电流正常负载测试示波器读取 = 100530,   3*2
	public static void genFunc100530(XWPFParagraph doc,TestResultData data)throws Exception{
		
		String base64 = data.getTestExportData().get("imageBase64").getAsString();
		XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
		
		ImgUtil.insertForPara(para,base64,400,400);
		XWPFTable table=POITableUtil.createCursorTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "灯丝电流有效值测试");
//
//		
		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("fCurrMin").getAsString()+"~"+ data.getTestInputData().get("fCurrMax").getAsString());
//
		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("CurrRMS").getAsDouble());
//
		POITableUtil.createCursorParagraph(doc);

//		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 2);
//		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
//		POITableUtil.setCellValueString(table, 0, 0, "");
//		POITableUtil.setCellValueString(table, 0, 1, "灯丝电流有效值测试");
//
////		XmlCursor cursor = table.getRow(0).getCell(0).addParagraph().getCTP().newCursor();
////		XWPFTable tableOne = doc.insertNewTbl(cursor);
//		
//		POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
//		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("fCurrMin").getAsString()+"~"+ data.getTestInputData().get("fCurrMax").getAsString());
//
//		POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
//		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("CurrRMS").getAsDouble());
//
//		doc.createParagraph();
	}




	//灯丝限流电阻更换 = 100550,
	public static void genFunc100550(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "R260*");
		POITableUtil.setCellValueString(table, 0, 2, "R264*");

		POITableUtil.setCellValueString(table, 1, 0, "更换前电阻值");
		POITableUtil.setCellValueDouble(table, 1, 1,  data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 2, 0, "更换后电阻值");
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//波形峰值测试D62  100570
	public static void genFunc100570(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 2, 4);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "高压开浪涌（A）");
		POITableUtil.setCellValueString(table, 0, 2, "阴极上升波形（V）");
		POITableUtil.setCellValueString(table, 0, 3, "D62)测试波形（V）");

		POITableUtil.setCellValueString(table, 1, 0, "实际测试值");
		POITableUtil.setCellValueDouble(table, 1, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 3,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//螺旋极调试 = 100610
	public static void genFunc100610(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 2, 5);
		POITableUtil.setTableColmnWidth(table, new int[]{1500,1500,1500,1500,1500});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
		POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
		POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
		POITableUtil.setCellValueString(table, 0, 4, "电阻值");

		POITableUtil.setCellValueString(table, 1, 0, "调试前");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueDouble(table, 1, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueString(table, 1, 3, "计算值");
		POITableUtil.setCellValueDouble(table, 1, 4,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//波形峰值测试全桥浪涌  100640
	public static void genFunc100640(XWPFDocument doc,TestResultData data)throws Exception{
		String base64 = data.getTestExportData().get("busImageBase64").getAsString();
		XWPFParagraph para = doc.createParagraph();
		ImgUtil.insertForPara(para,base64,400,400);
		doc.createParagraph();


		XWPFTable table=POITableUtil.createCommonTable(doc, 2, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "高压开浪涌（A）");
		POITableUtil.setCellValueString(table, 0, 2, "全桥浪涌（A）");

		POITableUtil.setCellValueString(table, 1, 0, "实际测试值");
		POITableUtil.setCellValueDouble(table, 1, 1,   data.getTestExportData().get("busWaveMax").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 2,  data.getTestExportData().get("llWaveMax").getAsDouble());
		doc.createParagraph();
	}

	//自动调阳压 = 100650,
	public static void genFunc100650(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 2, 5);
		POITableUtil.setTableColmnWidth(table, new int[]{1500,1500,1500,1500,1500});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
		POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
		POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
		POITableUtil.setCellValueString(table, 0, 4, "电阻值");

		POITableUtil.setCellValueString(table, 1, 0, "调试前");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueDouble(table, 1, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueString(table, 1, 3, "计算值");
		POITableUtil.setCellValueDouble(table, 1, 4,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//电源杂波 = 120050
	public static void genFunc120050(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "杂波频率（KHz）");
		POITableUtil.setCellValueDouble(table, 0, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 0, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 1, 0, "指标（dBc）");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueString(table, 1, 2, "接口单");

		POITableUtil.setCellValueString(table, 2, 0, "测试结果（dBc）");
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());

		doc.createParagraph();
	}

	//饱和输入输出 = 120110
	public static void genFunc120110(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 4, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "频率（GHz）");
		POITableUtil.setCellValueDouble(table, 0, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 1, 0, "饱和输入功率（dBm）");
		POITableUtil.setCellValueDouble(table, 1, 1,  data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 2, 0, "饱和输出功率（dBm）");
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 3, 0, "饱和输出功率（W）");
		POITableUtil.setCellValueDouble(table, 3, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//输入输出特性 = 120120
		public static void genFunc120120(XWPFDocument doc, TestResultData data){

			JsonArray yaoCeDianYa = data.getTestExportData().get("YaoCeDianYa").getAsJsonArray();
			int size = yaoCeDianYa.size();
			System.out.println("yaoCeDianYa:"+size);
			XWPFTable table= POITableUtil.createCommonTable(doc, size+1, 11);
			POITableUtil.setTableColmnWidth(table, new int[]{1500,700,700,700,700,700,700,700,700,700,700});
			//表头行
			POITableUtil.setCellValueString(table, 0, 0, "101g\n\r(Pin/Pinsat)dB");
			POITableUtil.setCellValueString(table, 0, 1, "母线电流(A)");
			POITableUtil.setCellValueString(table, 0, 2, "输出功率(W)");
			POITableUtil.setCellValueString(table, 0, 3, "螺流遥测(V)");
			POITableUtil.setCellValueString(table, 0, 4, "阳压遥测(V)");
			POITableUtil.setCellValueString(table, 0, 5, "开关机状态遥测(V)");
			POITableUtil.setCellValueString(table, 0, 6, "母线电流遥测（V）");
			POITableUtil.setCellValueString(table, 0, 7, "自动重启遥测（V）");
			POITableUtil.setCellValueString(table, 0, 8, "boost-buck电压(V)");
			POITableUtil.setCellValueString(table, 0, 9, "deltH电压值(V)");
			POITableUtil.setCellValueString(table, 0, 10, "阳极电压测试值(V)");


			/*for (int i = 1; i < yaoCeDianYa.size()+1; i++) {//每一行内容
				for(int j=0;j<11;j++) {                                       //MuXianDianLiu
					JsonArray muXianDianLiu = data.getTestExportData().get("MuXianDianLiu").getAsJsonArray();// 获取MuXianDianLiu里面的数据
					for (int i1 = 0; i1 < muXianDianLiu.size(); i1++) {
						POITableUtil.setCellValueString(table, i, j, data.getTestExportData().get("MuXianDianLiu").getAsJsonArray().getAsString());
					}
				}*/

		     //遥测电压
			for (int i = 0; i < yaoCeDianYa.size(); i++) {//每一行内容
				JsonObject jsonObject = yaoCeDianYa.get(i).getAsJsonObject();
				double luoLIuYaoCe = jsonObject.get("LuoLIuYaoCe").getAsDouble();
				double yangYaYaoCe = jsonObject.get("YangYaYaoCe").getAsDouble();
				double kaiGuanjiYaoCe = jsonObject.get("KaiGuanjiYaoCe").getAsDouble();
				double muXianDianLiuYaoCe = jsonObject.get("MuXianDianLiuYaoCe").getAsDouble();
				double ziDongChongQiYaoCe = jsonObject.get("ZiDongChongQiYaoCe").getAsDouble();
				//将yaoCeDianYa转换为map
				/*Gson gson = new Gson();
				Map map = gson.fromJson(yaoCeDianYa, Map.class);
				String luoLIuYaoCe = (String) map.get("LuoLIuYaoCe");//获取到的LuoLIuYaoCe
				String yangYaYaoCe = (String) map.get("YangYaYaoCe");//获取到的YangYaYaoCe
				String kaiGuanjiYaoCe = (String) map.get("KaiGuanjiYaoCe");//获取到的KaiGuanjiYaoCe
				String muXianDianLiuYaoCe = (String) map.get("MuXianDianLiuYaoCe");//获取到的MuXianDianLiuYaoCe
				String ziDongChongQiYaoCe = (String) map.get("ZiDongChongQiYaoCe");//获取到ZiDongChongQiYaoCe*/

				POITableUtil.setCellValueDouble(table, i+1, 3,luoLIuYaoCe);
				POITableUtil.setCellValueDouble(table, i+1, 4,yangYaYaoCe);
				POITableUtil.setCellValueDouble(table, i+1, 5,kaiGuanjiYaoCe);
				POITableUtil.setCellValueDouble(table, i+1, 6,muXianDianLiuYaoCe);
				POITableUtil.setCellValueDouble(table, i+1, 7,ziDongChongQiYaoCe);

			/*  POITableUtil.setCellValueString(table, i+1, 4, yaoCeDianYa.get("YangYaYaoCe").getAsString());
				POITableUtil.setCellValueString(table, i+1, 5, yaoCeDianYa.get("KaiGuanjiYaoCe").getAsString());
				POITableUtil.setCellValueString(table, i+1, 6, yaoCeDianYa.get("MuXianDianLiuYaoCe").getAsString());
				POITableUtil.setCellValueString(table, i+1, 7, yaoCeDianYa.get("ZiDongChongQiYaoCe").getAsString());
			*/
			}

			//母线电流
			JsonArray muXianDianLiu = data.getTestExportData().get("MuXianDianLiu").getAsJsonArray();
			int size1 = muXianDianLiu.size();
			System.out.println("MuXianDianLiu:"+size1);
			for (int i = 0; i < size1; i++) {//每一行内容
				JsonObject jsonObject = muXianDianLiu.get(i).getAsJsonObject();
				double curr = jsonObject.get("Curr").getAsDouble();
				POITableUtil.setCellValueDouble(table, i + 1, 1, curr);
			}

			//输出功率
			JsonArray ibOs = data.getTestExportData().get("IBOs").getAsJsonArray();
			int size2 = ibOs.size();
			System.out.println("IBOs:"+size2);
			for (int i = 0; i < size2; i++) {//每一行内容
				JsonObject jsonObject = ibOs.get(i).getAsJsonObject();
				double pOut = jsonObject.get("POut").getAsDouble();//输出功率
				POITableUtil.setCellValueDouble(table, i + 1, 2, pOut);

			}

            //BoostBuck
			JsonArray boostBuckYaoCe = data.getTestExportData().get("BoostBuckYaoCe").getAsJsonArray();
			int size3 = boostBuckYaoCe.size();
			System.out.println("BoostBuckYaoCe:"+size3);
			for (int i = 0; i < size3; i++) {//每一行内容
				JsonObject jsonObject = boostBuckYaoCe.get(i).getAsJsonObject();
				double boostBuckYaoCe1 = jsonObject.get("BoostBuckYaoCe").getAsDouble();

				POITableUtil.setCellValueDouble(table, i + 1, 8, boostBuckYaoCe1);
			}

			//DeltHs
			JsonArray deltHs = data.getTestExportData().get("DeltHs").getAsJsonArray();
			int size4 = deltHs.size();
			System.out.println("DeltHs:"+size4);
			for (int i = 0; i < size4; i++) {//每一行内容
				JsonObject jsonObject = deltHs.get(i).getAsJsonObject();
				double deltHVolt = jsonObject.get("DeltHVolt").getAsDouble();
				POITableUtil.setCellValueDouble(table, i + 1, 9, deltHVolt);
			}

			//阳极电压测试
			JsonArray ibOs1 = data.getTestExportData().get("IBOs").getAsJsonArray();
			int size5 = ibOs1.size();

			System.out.println("阳极电压测试:"+size5);
			for (int i = 0; i < size5; i++) {//每一行内容
				JsonObject jsonObject = ibOs1.get(i).getAsJsonObject();
				double voltA = jsonObject.get("VoltA").getAsDouble();//输出功率
				POITableUtil.setCellValueDouble(table, i + 1, 10, voltA);


			}
			for (int j = -20; j <= 3; j++) {
				try {
					XWPFParagraph pag = table.getRow(j + 21).getCell(0).addParagraph();
					pag.setAlignment(ParagraphAlignment.CENTER);
					pag.createRun().setText(j+"");
				}catch (Exception e){
					e.printStackTrace();
				}//POITableUtil.setCellValueDouble(table, j+ 21, 0, j);

			}


			doc.createParagraph();
			}




	//谐波测试 = 120130
	public static void genFunc120130(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{1000,1000,1000});
		POITableUtil.setCellValueString(table, 0, 0, "谐波次数");
		POITableUtil.setCellValueString(table, 0, 1, "2");
		POITableUtil.setCellValueString(table, 0, 2, "3");

		POITableUtil.setCellValueString(table, 1, 0, "指标(dBc)");
		POITableUtil.setCellValueString(table, 1, 1, data.getTestInputData().get("xbYiZhiDu").getAsString());
		POITableUtil.setCellValueString(table, 1, 2, data.getTestInputData().get("xbYiZhiDu").getAsString());

		POITableUtil.setCellValueString(table, 2, 0, "测试结果（dBc）");
		POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("XB_2_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}
	
	//带外杂波 = 120150
	public static void genFunc120150(XWPFDocument doc, TestResultData data){

		System.out.println(120150);
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 4);
		POITableUtil.setTableColmnWidth(table, new int[]{1500,1500,1500,1500});
		POITableUtil.setCellValueString(table, 0, 0, "带外杂波频率");
	/*	POITableUtil.setCellValueDouble(table, 0, 1, data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 0, 2,  data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 0, 3, data.getTestExportData().get("XB_3_Value").getAsDouble());*/


		POITableUtil.setCellValueString(table, 1, 0, "指标(dBc)");
		/*POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueString(table, 1, 2, "接口单");
		POITableUtil.setCellValueString(table, 1, 3, "接口单");*/

		POITableUtil.setCellValueString(table, 2, 0, "测试结果");
		//zbs
		JsonArray zbs = data.getTestExportData().get("zbs").getAsJsonArray();
		int size = zbs.size();
		System.out.println("zbs:"+size);
		if(size>3){
			size=3;
		}
		for (int i = 0; i < size; i++) {//每一行内容
			JsonObject jsonObject = zbs.get(i).getAsJsonObject();
			double freq = jsonObject.get("freq").getAsDouble();
			POITableUtil.setCellValueDouble(table, 0, i + 1,freq);

			double level = jsonObject.get("level").getAsDouble();
			POITableUtil.setCellValueDouble(table, 2, i + 1,level);

			POITableUtil.setCellValueDouble(table, 1, i + 1, data.getTestInputData().get("zhiBiao").getAsDouble());

		}






		doc.createParagraph();
	}



	//矢网推饱和 = 120310
	public static void genFunc120310(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 2, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "频率（GHz）");
		POITableUtil.setCellValueDouble(table, 0, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 1, 0, "饱和输入功率（dBm）");
		POITableUtil.setCellValueDouble(table, 1, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		doc.createParagraph();
	}

	//增益平坦度 = 120920
	public static void genFunc120920(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "项目");
		POITableUtil.setCellValueString(table, 0, 1, "饱和增益平坦度");
		POITableUtil.setCellValueString(table, 0, 2, "饱和增益平坦度斜率");

		POITableUtil.setCellValueString(table, 1, 0, "指标（dB）");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueString(table, 1, 2, "接口单");

		POITableUtil.setCellValueString(table, 2, 0, "测试值（dB）");
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		doc.createParagraph();
	}

	//群时延跳变 = 120930
	public static void genFunc120930(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 5, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "测试功率");
		POITableUtil.setCellValueString(table, 0, 1, "跳变幅值");

		POITableUtil.setCellValueDouble(table, 1, 0, data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueDouble(table, 2, 0,  data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueDouble(table, 3, 0,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 3, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueDouble(table, 4, 0,  data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 4, 1,  data.getTestExportData().get("XB_3_Value").getAsDouble());

		doc.createParagraph();
	}

	//相移_AMPM = 120940
	public static void genFunc120940(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 6);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,1500,1500,1500,1500,1500});
		POITableUtil.setCellValueString(table, 0, 0, "测试点");
		POITableUtil.setCellValueString(table, 0, 1, "接口单");
		POITableUtil.setCellValueString(table, 0, 2, "接口单");
		POITableUtil.setCellValueString(table, 0, 3, "接口单");
		POITableUtil.setCellValueString(table, 0, 4, "接口单");
		POITableUtil.setCellValueString(table, 0, 5, "接口单");

		POITableUtil.setCellValueString(table, 1, 0, "指标（DEG）");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueString(table, 1, 2, "接口单");
		POITableUtil.setCellValueString(table, 1, 3, "接口单");
		POITableUtil.setCellValueString(table, 1, 4, "接口单");
		POITableUtil.setCellValueString(table, 1, 5, "接口单");

		POITableUtil.setCellValueString(table, 2, 0, "测试结果（DEG）");
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 4,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 5,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}


	//带外增益 = 120960
	public static void genFunc120960(XWPFDocument doc,TestResultData data){
		XWPFTable table=POITableUtil.createCommonTable(doc, 3, 4);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,1000,1000,1000});
		POITableUtil.setCellValueString(table, 0, 0, "带外频率");
		POITableUtil.setCellValueDouble(table, 0, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 0, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 0, 3,   data.getTestExportData().get("XB_3_Value").getAsDouble());


		POITableUtil.setCellValueString(table, 1, 0, "指标（dB）");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueString(table, 1, 2, "接口单");
		POITableUtil.setCellValueString(table, 1, 3, "接口单");

		POITableUtil.setCellValueString(table, 2, 0, "测试结果（dB）");
		POITableUtil.setCellValueDouble(table, 2, 1,  data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 3,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		doc.createParagraph();
	}

	//************************************************************

	//低压开灯丝电压_电流连线提示 = 200120
	public static void genFunc200120(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 3);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "灯丝电压有效值（V）");
		POITableUtil.setCellValueString(table, 0, 2, "灯丝电流有效值（A）");

		POITableUtil.setCellValueString(table, 1, 0, "指标要求值");
		POITableUtil.setCellValueDouble(table, 1, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 2,  data.getTestExportData().get("XB_3_Value").getAsDouble());

		POITableUtil.setCellValueString(table, 2, 0, "实际测试值");
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//  提示并选择阳压截止状态高压负载阻值变化 = 200150
	public static void genFunc200150(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 3, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "阳压截止");
		POITableUtil.setCellValueString(table, 0, 1, "阳压导通");

		POITableUtil.setCellValueString(table, 1, 0, "接口单");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");

		POITableUtil.setCellValueDouble(table, 2, 0,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 2, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		doc.createParagraph();
	}

	//将负载箱阻值恢复为正常状态 = 200170
	public static void genFunc200170(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 2, 2);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "收集极");
		POITableUtil.setCellValueString(table, 0, 1, "阻抗");

		POITableUtil.setCellValueDouble(table, 1, 0,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueDouble(table, 1, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());

		doc.createParagraph();
	}

     //************************************************************

	//阳极测试结果输出 = 300020
	public static void genFunc300020(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 2, 4);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
		POITableUtil.setCellValueString(table, 0, 0, "标称值（V）");
		POITableUtil.setCellValueString(table, 0, 1, "测试值（V）");
		POITableUtil.setCellValueString(table, 0, 2, "实际值（V）");
		POITableUtil.setCellValueString(table, 0, 3, "差值（V）");

		POITableUtil.setCellValueString(table, 1, 0, "接口单");
		POITableUtil.setCellValueDouble(table, 1, 1,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueString(table, 1, 2, "计算值");
		POITableUtil.setCellValueString(table, 1, 3, "计算值");

		doc.createParagraph();
	}

	//判断阳极电压处于导通或截至状态 = 500070
	public static void genFunc500070(XWPFDocument doc, TestResultData data){
		XWPFTable table= POITableUtil.createCommonTable(doc, 2, 5);
		POITableUtil.setTableColmnWidth(table, new int[]{2000,2000,2000,2000,2000});

		POITableUtil.setCellValueString(table, 0, 0, "");
		POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
		POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
		POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
		POITableUtil.setCellValueString(table, 0, 4, "差值（V）");

		POITableUtil.setCellValueString(table, 1, 0, "调试前");
		POITableUtil.setCellValueString(table, 1, 1, "接口单");
		POITableUtil.setCellValueDouble(table, 1, 2,   data.getTestExportData().get("XB_3_Value").getAsDouble());
		POITableUtil.setCellValueString(table, 1, 3, "计算值");
		POITableUtil.setCellValueString(table, 1, 4, "计算值");
		doc.createParagraph();
	}





}
