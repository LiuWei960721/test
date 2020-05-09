package com.ss.poi.service;

import com.ss.poi.util.*;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.ss.poi.entity.EPCVoltStageEnum;
import com.ss.poi.entity.TestResultData;

import java.math.BigDecimal;

/**
 * @Description
 */
public class GroupCA {

    // 母线电流100030
    public static void genFunc100030(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1,
                data.getEPCStage().getVoltStageStr() + data.getEPCStage().getBusVoltStr() + "电流值(A)");
        POITableUtil.setCellValueString(table, 0, 2, "母线功耗(W)");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值");
        // 接口单
        switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getEPCVoltStage())) {
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

        switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getEPCVoltStage())) {
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
        POITableUtil.createCursorParagraph(doc);
    }

    // 遥测电压100040
    public static void genFunc100040(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 6);
        POITableUtil.setTableColmnWidth(table, new int[]{1900, 1200, 1200, 1200, 1200, 1200});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "螺流遥测");
        POITableUtil.setCellValueString(table, 0, 2, "阳压遥测");
        POITableUtil.setCellValueString(table, 0, 3, "开关机状态遥测");
        POITableUtil.setCellValueString(table, 0, 4, "自动重启遥测");
        POITableUtil.setCellValueString(table, 0, 5, "母线电流遥测");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");// yanYaYaoCeVoltMin
        POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("luoLiuYaoCeVoltMin").getAsDouble())
                + "~" + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("luoLiuYaoCeVoltMax").getAsDouble()));
        POITableUtil.setCellValueString(table, 1, 2,
                DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("yanYaYaoCeVoltMin").getAsDouble()) + "~" + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("yanYaYaoCeVoltMax").getAsDouble()));
        POITableUtil.setCellValueString(table, 1, 3, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("kaiGuanJiYaoCeVoltMin").getAsDouble())
                + "~" + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("kaiGuanJiYaoCeVoltMax").getAsDouble()));
        POITableUtil.setCellValueString(table, 1, 4,
                DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("ziDongChongQiYaoCeVoltMin").getAsDouble()) + "~"
                        + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("ziDongChongQiYaoCeVoltMax").getAsDouble()));
        POITableUtil.setCellValueString(table, 1, 5,
                DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("muXianYaoCeVoltMin").getAsDouble()) + "~" + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("muXianYaoCeVoltMin").getAsDouble()));

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("LuoLIuYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("YangYaYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 3, data.getTestExportData().get("KaiGuanjiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 4, data.getTestExportData().get("ZiDongChongQiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 5, data.getTestExportData().get("MuXianDianLiuYaoCe").getAsDouble());
        POITableUtil.createCursorParagraph(doc);

    }

    // BoostBuck电压测试100050
    public static void genFunc100050(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2500});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "boost-buck电压测试值");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
        POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("BoostBuckMin").getAsDouble()) + "~"
                + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("BoostBuckMax").getAsDouble()));

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("BoostBuckYaoCe").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // CAMP电压电流测试 = 100060
    public static void genFunc100060(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 5, 4);
        POITableUtil.setTableColmnWidth(table, new int[]{3000, 1500, 1500, 1500});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "Camp+5");
        POITableUtil.setCellValueString(table, 0, 2, "Camp+12");
        POITableUtil.setCellValueString(table, 0, 3, "Camp-12");
        /*改动了类型*/
        POITableUtil.setCellValueString(table, 1, 0, "指标电压要求值（V）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("zhengV5V").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestInputData().get("zhengV12V").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 3, data.getTestInputData().get("fuV12V").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "指标电流要求值（A）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestInputData().get("zhengI5V").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestInputData().get("zhengI12V").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 3, data.getTestInputData().get("fuI12V").getAsDouble());

        POITableUtil.setCellValueString(table, 3, 0, "实际电压测试值（V）");
        POITableUtil.setCellValueDouble(table, 3, 1, data.getTestExportData().get("VZheng5").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 2, data.getTestExportData().get("VZheng12").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 3, data.getTestExportData().get("VFu12").getAsDouble());

        POITableUtil.setCellValueString(table, 4, 0, "实际电流测试值（A）");
        POITableUtil.setCellValueDouble(table, 4, 1, data.getTestExportData().get("IZheng5").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 2, data.getTestExportData().get("IZheng12").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 3, data.getTestExportData().get("IFu12").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // deltH电压测试 = 100070
    public static void genFunc100070(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "deltH电压测试值");
//修改
        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
        POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("deltHMin").getAsDouble()) + "~"
                + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("deltHMax").getAsDouble()));

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("DeltHVolt").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // HC1C2 100080
    public static void genFunc100080(XWPFParagraph doc, TestResultData data) {

        EPCVoltStageEnum voltStage = data.getEPCStage().getEPCVoltStageEnum();
        switch (voltStage) {
            case 静态:
            case 动态:
                genFunc100080B(doc, data);
                break;
            case 低压:
            default:
                genFunc100080A(doc, data);
                return;
        }
    }

    public static void genFunc100080A(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 9, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 1000, 1000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
        POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");

        POITableUtil.setCellValueString(table, 1, 0, "螺旋级（H）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("setVoltH").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("testVoltH").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "阳极（A）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("setVoltA1").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("testVoltA1").getAsDouble());

        POITableUtil.setCellValueString(table, 3, 0, "收集极（C1）");
        POITableUtil.setCellValueDouble(table, 3, 1, data.getTestExportData().get("setVoltC1").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 2, data.getTestExportData().get("testVoltC1").getAsDouble());

        POITableUtil.setCellValueString(table, 4, 0, "收集极（C2）");
        POITableUtil.setCellValueDouble(table, 4, 1, data.getTestExportData().get("setVoltC2").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 2, data.getTestExportData().get("testVoltC2").getAsDouble());

        POITableUtil.setCellValueString(table, 5, 0, "收集极（C3）");
        POITableUtil.setCellValueDouble(table, 5, 1, data.getTestExportData().get("setVoltC3").getAsDouble());
        POITableUtil.setCellValueDouble(table, 5, 2, data.getTestExportData().get("testVoltC3").getAsDouble());

        POITableUtil.setCellValueString(table, 6, 0, "收集极（C4）");
        POITableUtil.setCellValueDouble(table, 6, 1, data.getTestExportData().get("setVoltC4").getAsDouble());
        POITableUtil.setCellValueDouble(table, 6, 2, data.getTestExportData().get("testVoltC4").getAsDouble());

        POITableUtil.setCellValueString(table, 7, 0, "阴极（K）");
        POITableUtil.setCellValueDouble(table, 7, 1, data.getTestExportData().get("setVoltK").getAsDouble());
        POITableUtil.setCellValueDouble(table, 7, 2, data.getTestExportData().get("testVoltK").getAsDouble());

        POITableUtil.setCellValueString(table, 8, 0, "A2");
        POITableUtil.setCellValueString(table, 8, 1, data.getTestExportData().get("setVoltA2").getAsString());
        POITableUtil.setCellValueDouble(table, 8, 2, data.getTestExportData().get("testVoltA2").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    public static void genFunc100080B(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 9, 5);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 1000, 1000, 1000, 1000});
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

        POITableUtil.setCellValueString(table, 8, 0, "A2");
        POITableUtil.setCellValueString(table, 8, 1, data.getTestExportData().get("setVoltA2").getAsString());
        POITableUtil.setCellValueDouble(table, 8, 2, data.getTestExportData().get("testVoltA2").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 电阻电容检查与设置 = 100130
    public static void genFunc100130(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 6, 4);
        POITableUtil.setTableColmnWidth(table, new int[]{1000, 1000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "序号");
        POITableUtil.setCellValueString(table, 0, 1, "");
        POITableUtil.setCellValueString(table, 0, 2, "");
        POITableUtil.setCellValueString(table, 0, 3, "程序回读值");

        POITableUtil.setCellValueString(table, 1, 0, "1");

        JsonObject inData = data.getTestInputData();
        JsonObject exData = data.getTestExportData();

        switch (data.getFuncType()) {
            case 100130:
                POITableUtil.setCellValueString(table, 1, 1, "PLC初始安装");
                POITableUtil.setCellValueDouble(table, 3, 3, inData.get("resLuoxuanji").getAsDouble());
                break;
            case 300030:
                POITableUtil.setCellValueString(table, 1, 1, "PLC最终确定值");
                POITableUtil.setCellValueDouble(table, 3, 3, exData.get("YinjiKR").getAsDouble());
                break;
        }

        POITableUtil.setCellValueString(table, 1, 2, "");
        POITableUtil.setCellValueString(table, 1, 3, "");

        POITableUtil.setCellValueString(table, 2, 0, "");
        POITableUtil.setCellValueString(table, 2, 1, "");
        POITableUtil.setCellValueString(table, 2, 2, "boost-buck电阻值");
        POITableUtil.setCellValueDouble(table, 2, 3, data.getTestExportData().get("BoostBuckR").getAsDouble());

        POITableUtil.setCellValueString(table, 3, 0, "");
        POITableUtil.setCellValueString(table, 3, 1, "");
        POITableUtil.setCellValueString(table, 3, 2, "K电阻");

        /*
         * POITableUtil.setCellValueString(table, 4, 0, "");
         * POITableUtil.setCellValueString(table, 4, 1, "");
         * POITableUtil.setCellValueString(table, 4, 2,"R6R7电阻值");
         * POITableUtil.setCellValueDouble(table, 4, 3,
         * data.getTestExportData().get("BoostBuckR").getAsDouble());
         */
        POITableUtil.setCellValueString(table, 4, 0, "");
        POITableUtil.setCellValueString(table, 4, 1, "");
        POITableUtil.setCellValueString(table, 4, 2, "阳极A电阻");
        POITableUtil.setCellValueDouble(table, 4, 3, exData.get("resYangYa").getAsDouble());

        POITableUtil.setCellValueString(table, 5, 0, "");
        POITableUtil.setCellValueString(table, 5, 1, "");
        POITableUtil.setCellValueString(table, 5, 2, "电容值");
        POITableUtil.setCellValueDouble(table, 5, 3, data.getTestExportData().get("XieZhengDianRongR").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 低压开BoostBusk电压测试 100150
    public static void genFunc100150(XWPFParagraph para, TestResultData data) throws Exception {
        if (para == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(para, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "调试前Boost-Buck电压");
        POITableUtil.setCellValueString(table, 0, 2, "调试后Boost-Buck电压");

        JsonObject exData = data.getTestExportData();
        POITableUtil.setCellValueString(table, 1, 0, "电阻值（Ω）");
        POITableUtil.setCellValueDouble(table, 1, 1, exData.get("BoostBuckTiaoShiQianDianZu").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, exData.get("BoostBuckTiaoShiHouDianZu").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "电压值（V）");
        POITableUtil.setCellValueDouble(table, 2, 1, exData.get("BoostBuckTiaoShiQianDianYa").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, exData.get("BoostBuckTiaoShiHouDianYa").getAsDouble());
        POITableUtil.createCursorParagraph(para);
        String imgPath = LineChartUtil.createChart(data.getTestExportData(), "电阻/Ω", "电压/V", "低压开BoostBusk电压测试");
        System.out.println("生成折线图：" + imgPath);
        XWPFParagraph imgPara = POITableUtil.createCursorParagraph(para);
        ImgUtil.insertForPara(imgPara, imgPath, 700, 400);
        POITableUtil.createCursorParagraph(para);
        XWPFParagraph picNextPara = POITableUtil.createCursorParagraph(para);
        OperateTable.createLineChartTable(picNextPara, data.getTestExportData());
        POITableUtil.createCursorParagraph(para);
    }

    // 回读电容 = 100180
    public static void genFunc100180(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "谐振电容值");

        POITableUtil.setCellValueString(table, 1, 0, "程序回读值");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("Capacitance").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 高压开浪涌 100190
    public static void genFunc100190(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);
        POITableUtil.createCursorParagraph(doc);

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");

        // 接口单
        switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getEPCVoltStage())) {
            case 低压:
                POITableUtil.setCellValueString(table, 0, 1, "低压开浪涌");
                break;
            case 动态:
            case 静态:
                POITableUtil.setCellValueString(table, 0, 1, "高压开浪涌");
                break;
            default:
                break;
        }
//修改
        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("turnonSurge").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("turnonSurge").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 环路稳定性测试 = 100200 低压
    public static void genFunc100200(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        switch (EPCVoltStageEnum.valueOf(data.getEPCStage().getEPCLoadStage())) {
            case 静态:
            case 动态:
                genFunc100200B(doc, data);
                break;
            case 低压:
            default:
                genFunc100200B(doc, data);
                return;
        }
    }

    // 环路稳定性测试 = 100200 低压
    public static void genFunc100200A(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "低压欠压保护值");
        POITableUtil.setCellValueString(table, 0, 2, "低压过压保护值");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
        POITableUtil.setCellValueString(table, 1, 1, "接口单");
        POITableUtil.setCellValueString(table, 1, 2, "接口单");

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
        POITableUtil.setCellValueString(table, 2, 1, data.getTestExportData().get("VoltProtect").getAsString());
        POITableUtil.setCellValueString(table, 2, 2, data.getTestExportData().get("VoltProtect").getAsString());
        POITableUtil.createCursorParagraph(doc);
    }

    // 环路稳定性测试 = 100200 高压
    public static void genFunc100200B(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "高压环路稳定性测试");

        POITableUtil.setCellValueString(table, 1, 0, "稳定性判断");
        POITableUtil.setCellValueString(table, 1, 1, "环路正常");

        POITableUtil.createCursorParagraph(doc);
    }

    // //环路稳定性测试 = 100200 高压
    // public static void genFunc100220B(XWPFParagraph doc,TestResultData data){
    // XWPFTable table=POITableUtil.createCursorTable(doc, 2, 2);
    // POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
    // POITableUtil.setCellValueString(table, 0, 0, "");
    // POITableUtil.setCellValueString(table, 0, 1, "高压环路稳定性测试");
    //
    // POITableUtil.setCellValueString(table, 1, 0, "稳定性判断");
    // POITableUtil.setCellValueString(table, 1, 1, "环路正常");
    //
    // POITableUtil.createCursorParagraph(doc);
    // }

    // 开机阈值测试 = 100230
    public static void genFunc100230(XWPFParagraph doc, TestResultData data) {
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        if (doc == null) {
            return;
        }
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "低压开机阈值");
//修改
        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("busVoltProtMax").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("VoltProtect").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 谐振波形 100240
    public static void genFunc100240(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);
        POITableUtil.createCursorParagraph(doc);

    }

    // 低压开栅极电压测试 100260
    public static void genFunc100260(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 1, "低压开栅极电压测试");
//修改
        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（V）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("shanjiV").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（V）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("gridVoltage").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 谐振电容设置 100270
    public static void genFunc100270(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 1, "当前设置值");

        POITableUtil.setCellValueString(table, 1, 0, "谐振电容值设置（nF）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("XiezhenDianrong").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 高压开栅极电压测试 100280
    public static void genFunc100280(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        XWPFTable table = POITableUtil.createCursorTable(doc, 4, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 1, "测试值(V)");

        JsonObject exData = data.getTestExportData();
        POITableUtil.setCellValueString(table, 1, 0, "阴极(K)");
        POITableUtil.setCellValueDouble(table, 1, 1, exData.get("testK").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "栅极(G)");
        POITableUtil.setCellValueDouble(table, 2, 1, exData.get("testG").getAsDouble());

        POITableUtil.setCellValueString(table, 3, 0, "计算值");
        POITableUtil.setCellValueDouble(table, 3, 1, exData.get("actG").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 改变负载箱导通截至阻值 = 100290
    public static void genFunc100290(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{3000, 3000, 3000});
        POITableUtil.setCellValueString(table, 0, 0, "收集极");
        POITableUtil.setCellValueString(table, 0, 1, "阻抗");
        POITableUtil.setCellValueString(table, 0, 2, "系数");
        /*修改*/
        POITableUtil.setCellValueDouble(table, 1, 0, data.getTestInputData().get("MaxCaption").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("SetRes").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("minCurr").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 低压开灯丝电流正常负载测试示波器读取 = 100530, 3*2
    public static void genFunc100530(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);

        ImgUtil.insertForPara(para, base64, 400, 300);

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "灯丝电流有效值测试灯丝电流有效值测试");
        //
        //修改
        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
        POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("fCurrMin").getAsDouble()) + "~"
                + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("fCurrMax").getAsDouble()));
        //
        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("CurrRMS").getAsDouble(), 3);
        //
        POITableUtil.createCursorParagraph(doc);

        // XWPFTable table=POITableUtil.createCursorTable(doc, 3, 2);
        // POITableUtil.setTableColmnWidth(table, new int[]{2000,2000});
        // POITableUtil.setCellValueString(table, 0, 0, "");
        // POITableUtil.setCellValueString(table, 0, 1, "灯丝电流有效值测试");
        //
        //// XmlCursor cursor =
        // table.getRow(0).getCell(0).addParagraph().getCTP().newCursor();
        //// XWPFTable tableOne = doc.insertNewTbl(cursor);
        //
        // POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
        // POITableUtil.setCellValueString(table, 1, 1,
        // data.getTestInputData().get("fCurrMin").getAsString()+"~"+
        // data.getTestInputData().get("fCurrMax").getAsString());
        //
        // POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
        // POITableUtil.setCellValueDouble(table, 2, 1,
        // data.getTestExportData().get("CurrRMS").getAsDouble());
        //
        // POITableUtil.createCursorParagraph(doc);
    }

    // 灯丝限流电阻更换 = 100550 缺少data文件
    public static void genFunc100550(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "R260*");
        POITableUtil.setCellValueString(table, 0, 2, "R264*");

        POITableUtil.setCellValueString(table, 1, 0, "更换前电阻值");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "更换后电阻值");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // 波形峰值测试D62 100570
    public static void genFunc100570(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);

        POITableUtil.createCursorParagraph(doc);

        XWPFTable table = null;
        JsonElement turnonSurge1;
        JsonElement turnonSurge2;
        JsonElement turnonSurge3;

        table = POITableUtil.createCursorTable(doc, 2, 1);
        /* POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});*/
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "高压开浪涌（A）");
        /*POITableUtil.setCellValueString(table, 0, 2, "D62测试波形（V）");*/

        turnonSurge1 = data.testExportData.getAsJsonObject("busWaveData").get("turnonSurge");
        POITableUtil.setCellValueDouble(table, 1, 0, turnonSurge1.getAsDouble());
       /* turnonSurge3 = data.testExportData.get("d62WaveDataLowVMax");
        POITableUtil.setCellValueDouble(table, 1, 2, turnonSurge3.getAsDouble());*/

        POITableUtil.createCursorParagraph(doc);
        /*table = POITableUtil.createCursorTable(doc, 4, 2);*/
        table = POITableUtil.createCursorTable(doc, 4, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 1, "阴极上升");
        POITableUtil.setCellValueString(table, 0, 2, "D62测试");
        POITableUtil.setCellValueString(table, 1, 0, "上升时间（ms）");
        POITableUtil.setCellValueString(table, 2, 0, "上升最大幅值（mv）");
        POITableUtil.setCellValueString(table, 3, 0, "上升稳定幅值(mv)");

        POITableUtil.setCellValueDouble(table, 1, 1,
                data.getTestExportData().get("YinJiShangShengShiJian").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 1,
                data.getTestExportData().get("ShangShengZuiDaFuZhi").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 1,
                data.getTestExportData().get("ShangShengWenDingFuZhi").getAsDouble());
        //D62的值
        POITableUtil.setCellValueDouble(table, 1, 2,
                data.getTestExportData().get("D62ShangShengShiJian").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2,
                data.getTestExportData().get("D62ShangShengZuiDaFuZhi").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 2,
                data.getTestExportData().get("D62ShangShengWenDingFuZhi").getAsDouble());

        // turnonSurge2 =
        // data.testExportData.getAsJsonObject("kWaveData").get("turnonSurge");
        // POITableUtil.setCellValueDouble(table, 1, 2, turnonSurge2.getAsDouble());

        // POITableUtil.setCellValueString(table, 0, 0, "");

        // POITableUtil.setCellValueString(table, 1, 0, "实际测试值");
        POITableUtil.createCursorParagraph(doc);
    }

    // 102080
    public static void genFunc102080(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);

        POITableUtil.createCursorParagraph(doc);

        XWPFTable table = null;
        JsonElement turnonSurge1;
        JsonElement turnonSurge2;
        JsonElement turnonSurge3;

        table = POITableUtil.createCursorTable(doc, 2, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 1, "高压开浪涌（A）");
        POITableUtil.setCellValueString(table, 0, 2, "D62波形（V）");

        turnonSurge1 = data.testExportData.getAsJsonObject("busWaveData").get("turnonSurge");
        POITableUtil.setCellValueDouble(table, 1, 1, turnonSurge1.getAsDouble());
        turnonSurge2 = data.testExportData.get("d62WaveDataLowVMax");
        POITableUtil.setCellValueDouble(table, 1, 2, turnonSurge2.getAsDouble());
        // POITableUtil.setCellValueString(table, 0, 0, "");

        // POITableUtil.setCellValueString(table, 1, 0, "实际测试值");
        POITableUtil.createCursorParagraph(doc);
    }

    // 高压开遥测阻抗 102090
    public static void genFunc102090(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 5, 6);
        POITableUtil.setTableColmnWidth(table, new int[]{1000, 1000, 1000, 1000, 1000, 1000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "螺流遥测");
        POITableUtil.setCellValueString(table, 0, 2, "阳压遥测");
        POITableUtil.setCellValueString(table, 0, 3, "开关机状态遥测");
        POITableUtil.setCellValueString(table, 0, 4, "自动重启遥测");
        POITableUtil.setCellValueString(table, 0, 5, "母线电流遥测");

        JsonObject inData = data.getTestInputData();
        JsonObject exData = data.getTestExportData();
        JsonObject bingruQian = exData.get("BingRuDianZuQian").getAsJsonObject();
        JsonObject bingruHou = exData.get("BingRuDianZuHou").getAsJsonObject();
        JsonObject jisuan = exData.get("JiSuan").getAsJsonObject();

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值“≤”(KΩ)");
        POITableUtil.setCellValueString(table, 1, 1, inData.get("luoLiuYaoCeDongTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 2, inData.get("yanYaYaoDongTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 3, inData.get("kaiGuanJiYaoCeDongTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 4, inData.get("ziDongChongQiYaoCeDongTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 5, inData.get("muXianYaoCeDongTaiRes").getAsString());

        POITableUtil.setCellValueString(table, 2, 0, "静态遥测电压值(V)");
        POITableUtil.setCellValueDouble(table, 2, 1, bingruQian.get("LuoLIuYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, bingruQian.get("YangYaYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 3, bingruQian.get("KaiGuanjiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 4, bingruQian.get("ZiDongChongQiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 5, bingruQian.get("MuXianDianLiuYaoCe").getAsDouble());

        POITableUtil.setCellValueString(table, 3, 0, "并联1遥测电压(V)");
        POITableUtil.setCellValueDouble(table, 3, 1, bingruHou.get("LuoLIuYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 2, bingruHou.get("YangYaYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 3, bingruHou.get("KaiGuanjiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 4, bingruHou.get("ZiDongChongQiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 3, 5, bingruHou.get("MuXianDianLiuYaoCe").getAsDouble());

        POITableUtil.setCellValueString(table, 4, 0, "实际计算值(KΩ)");
        POITableUtil.setCellValueDouble(table, 4, 1, jisuan.get("LuoLIuYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 2, jisuan.get("YangYaYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 3, jisuan.get("KaiGuanjiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 4, jisuan.get("ZiDongChongQiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 4, 5, jisuan.get("MuXianDianLiuYaoCe").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // 螺旋极调试 = 100610
    public static void genFunc100610(XWPFParagraph para, TestResultData data) throws Exception {
        if (para == null) {
            return;
        }
        if (data.getTestExportData().get("ShiFouXuYaoTiaoShi").getAsBoolean()) {
            XWPFTable table = POITableUtil.createCursorTable(para, 3, 5);
            POITableUtil.setTableColmnWidth(table, new int[]{1500, 1500, 1500, 1500, 1500});
            POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
            POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
            POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
            POITableUtil.setCellValueString(table, 0, 4, "电阻值");

            POITableUtil.setCellValueString(table, 1, 0, "调试前");
            POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("kVoltValue").getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, 2,
                    data.getTestExportData().get("TiaoShiQianCeShiZhi").getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, 3,
                    data.getTestExportData().get("TiaoShiQianShiJiZhi").getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, 4,
                    data.getTestExportData().get("TiaoShiQianDianZuZhi").getAsDouble());

            POITableUtil.setCellValueString(table, 2, 0, "调试后");
            POITableUtil.setCellValueDouble(table, 2, 1, data.getTestInputData().get("kVoltValue").getAsDouble());
            POITableUtil.setCellValueDouble(table, 2, 2,
                    data.getTestExportData().get("TiaoShiHouCeShiZhi").getAsDouble());
            POITableUtil.setCellValueDouble(table, 2, 3,
                    data.getTestExportData().get("TiaoShiHouShiJiZhi").getAsDouble());
            POITableUtil.setCellValueDouble(table, 2, 4,
                    data.getTestExportData().get("TiaoShiHouDianZuZhi").getAsDouble());
            POITableUtil.createCursorParagraph(para);
            String imgPath = LineChartUtil.createChart(data.getTestExportData(), "电阻/Ω", "电压/V", "螺旋极调试");
            System.out.println("生成折线图：" + imgPath);
            XWPFParagraph imgPara = POITableUtil.createCursorParagraph(para);
            ImgUtil.insertForPara(imgPara, imgPath, 700, 400);
            POITableUtil.createCursorParagraph(para);
            XWPFParagraph picNextPara = POITableUtil.createCursorParagraph(para);
            OperateTable.createLineChartTable(picNextPara, data.getTestExportData());
            POITableUtil.createCursorParagraph(para);
            // ImgUtil.createLineImgHasXAxisNoLabel(listValue);

        } else {
            para.createRun().setText("电阻满足要求，不需要调试");
            XWPFTable table = POITableUtil.createCursorTable(para, 2, 5);
            POITableUtil.setTableColmnWidth(table, new int[]{1500, 1500, 1500, 1500, 1500});
            POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
            POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
            POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
            POITableUtil.setCellValueString(table, 0, 4, "电阻值");

            POITableUtil.setCellValueString(table, 1, 0, "调试前");
            POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("kVoltValue").getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, 2,
                    data.getTestExportData().get("TiaoShiQianCeShiZhi").getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, 3,
                    data.getTestExportData().get("TiaoShiQianShiJiZhi").getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, 4,
                    data.getTestExportData().get("TiaoShiQianDianZuZhi").getAsDouble());

        }
        POITableUtil.createCursorParagraph(para);
    }

    // 阳压A上升时间测试
    public static void genFunc100630(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);

        XWPFTable table = POITableUtil.createCursorTable(doc, 1, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "上升时间(ms)");

        POITableUtil.setCellValueDouble(table, 0, 1, data.getTestExportData().get("aWenDingShiJian").getAsDouble() * 1000d, 2);

        POITableUtil.createCursorParagraph(doc);
    }

    // 阴极K上升时间测试 = 100620
    public static void genFunc100620(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "上升时间(ms)");
        POITableUtil.setCellValueString(table, 1, 0, "上升最大幅值(V)");
        POITableUtil.setCellValueString(table, 2, 0, "上升稳定幅值(V)");

        POITableUtil.setCellValueDouble(table, 0, 1,
                data.getTestExportData().get("YinJiShangShengShiJian").getAsDouble() * 1000d, 2);
        POITableUtil.setCellValueDouble(table, 1, 1,
                data.getTestExportData().get("ShangShengZuiDaFuZhi").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 1,
                data.getTestExportData().get("ShangShengWenDingFuZhi").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 波形峰值测试全桥浪涌 100640
    public static void genFunc100640(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        String base64 = data.getTestExportData().get("busImageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);
        POITableUtil.createCursorParagraph(doc);

        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "高压开浪涌（A）");
        POITableUtil.setCellValueString(table, 0, 2, "全桥浪涌（A）");

        POITableUtil.setCellValueString(table, 1, 0, "实际测试值");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("busWaveMax").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("llWaveMax").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // 自动调阳压 = 100650
    public static void genFunc100650(XWPFParagraph para, TestResultData data) throws Exception {
        if (para == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(para, 2, 5);
        POITableUtil.setTableColmnWidth(table, new int[]{1500, 1500, 1500, 1500, 1500});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "标称值（V）");
        POITableUtil.setCellValueString(table, 0, 2, "测试值（V）");
        POITableUtil.setCellValueString(table, 0, 3, "实际值（V）");
        POITableUtil.setCellValueString(table, 0, 4, "电阻值");

        POITableUtil.setCellValueString(table, 1, 0, "调试前");
        POITableUtil.setCellValueString(table, 1, 1, "接口单");
        //POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueString(table, 1, 3, "计算值");
        //POITableUtil.setCellValueDouble(table, 1, 4, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.createCursorParagraph(para);
        String imgPath = LineChartUtil.createChart(data.getTestExportData(), "电阻/Ω", "电压/V", "自动调阳压");
        System.out.println("生成折线图：" + imgPath);
        XWPFParagraph imgPara = POITableUtil.createCursorParagraph(para);
        ImgUtil.insertForPara(imgPara, imgPath, 500, 300);
        POITableUtil.createCursorParagraph(para);
        XWPFParagraph picNextPara = POITableUtil.createCursorParagraph(para);
        OperateTable.createLineChartTable(picNextPara, data.getTestExportData());
        POITableUtil.createCursorParagraph(para);
    }

    // 阳极导通及截止测试 = 100660,
    public static void genFunc100660(XWPFParagraph doc, TestResultData data) {
        // XWPFTable table = POITableUtil.createCursorTable(doc, 2, 2);
        // POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        // POITableUtil.setCellValueString(table, 0, 0, "");
        // JsonElement caption =
        // data.testInputData.getAsJsonObject("EPCStage").get("Caption");
        // POITableUtil.setCellValueString(table, 0, 1, caption.getAsString());

        // POITableUtil.setCellValueString(table, 1, 0, "测试结果");
        // POITableUtil.setCellValueDouble(table, 1, 1,
        // data.getTestExportData().get("ASwithRes").getAsDouble());

        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        String msg = "阳极导通及截止测试结果：" + data.getTestExportData().get("ASwithRes").getAsString();
        para.createRun().setText(msg);
        POITableUtil.createCursorParagraph(doc);
    }

    // 灯丝预热时间测试 = 100710,
    public static void genFunc100710(XWPFParagraph doc, TestResultData data) {
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "预热时间");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求（s）");
        POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("dengsiMinTime").getAsDouble()) + "~"
                + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("dengsiMaxTime").getAsDouble()));

        POITableUtil.setCellValueString(table, 2, 0, "实际值（s）");
        POITableUtil.setCellValueDouble(table, 2, 1,
                data.getTestExportData().get("DengSiYuReShiJian").getAsDouble() / 1000);
        POITableUtil.createCursorParagraph(doc);
    }

    // 阳压导通电阻设置 = 100810,
    // 阳压截止电阻设置 = 100820,
    public static void genFunc100810(XWPFParagraph doc, TestResultData data) {
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 1);
        POITableUtil.setTableColmnWidth(table, new int[]{3500});
        switch (data.getFuncType()) {
            case 100810:
                POITableUtil.setCellValueString(table, 0, 0, "导通电阻值");
                break;
            case 100820:
                POITableUtil.setCellValueString(table, 0, 0, "截止电阻值");
                break;
        }
        POITableUtil.setCellValueDouble(table, 1, 0,
                data.getTestExportData().get("YangYaDianZuSheZhiZhi").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 电源杂波 = 120050 缺少指标值
    public static void genFunc120050(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "杂波频率（KHz）");
        POITableUtil.setCellValueDouble(table, 0, 1, data.getTestExportData().get("xb30kFreq").getAsDouble());
        POITableUtil.setCellValueDouble(table, 0, 2, data.getTestExportData().get("xb60kFreq").getAsDouble());

        POITableUtil.setCellValueString(table, 1, 0, "指标（dBc）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("dyzbMax").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestInputData().get("dyzbMax").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "测试结果（dBc）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("xb30kLevel").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("xb60kLevel").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 饱和输入输出 = 120110 缺少频率、计算饱和输出功率
    public static void genFunc120110(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 4, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "频率（GHz）");
        POITableUtil.setCellValueDouble(table, 0, 1, data.getTestExportData().get("FreqGHz").getAsDouble());

        POITableUtil.setCellValueString(table, 1, 0, "饱和输入功率（dBm）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("PIn").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "饱和输出功率（dBm）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("POut").getAsDouble());

        POITableUtil.setCellValueString(table, 3, 0, "饱和输出功率（W）");
        POITableUtil.setCellValueDouble(table, 3, 1, data.getTestExportData().get("POut_W").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // 输入输出特性 = 120120
    public static void genFunc120120(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }

        int funcType = data.getFuncType();

        // 遥测电压
        JsonArray yaoCeDianYa = data.getTestExportData().get("YaoCeDianYa").getAsJsonArray();
        // 母线电流
        JsonArray muXianDianLiu = data.getTestExportData().get("MuXianDianLiu").getAsJsonArray();
        // 输出功率
        JsonArray ibOs = data.getTestExportData().get("IBOs").getAsJsonArray();
        // BoostBuck
        JsonArray boostBuckYaoCe = data.getTestExportData().get("BoostBuckYaoCe").getAsJsonArray();
        // DeltHs
        JsonArray deltHs = data.getTestExportData().get("DeltHs").getAsJsonArray();
        // 阳极电压测试
        JsonArray voltAs = data.getTestExportData().get("IBOs").getAsJsonArray();

        int size = yaoCeDianYa.size();
        System.out.println("yaoCeDianYa:" + size);
        XWPFTable table = null;
        if (funcType == 120120) {
            table = POITableUtil.createCursorTable(doc, size + 1, 12);
            POITableUtil.setTableColmnWidth(table,
                    new int[]{1500, 700, 700, 700, 700, 700, 700, 700, 700, 700, 700, 700});

            POITableUtil.setCellValueString(table, 0, 9, "boost-buck电压(V)");
            POITableUtil.setCellValueString(table, 0, 10, "deltH电压值(V)");
            POITableUtil.setCellValueString(table, 0, 11, "阳极电压测试值(V)");

        } else {
            table = POITableUtil.createCursorTable(doc, size + 1, 9);
            POITableUtil.setTableColmnWidth(table, new int[]{1500, 700, 700, 700, 700, 700, 700, 700, 700});
        }

        // 表头行
        POITableUtil.setCellValueString(table, 0, 0, "101g\n\r(Pin/Pinsat)dB");
        POITableUtil.setCellValueString(table, 0, 1, "母线电流(A)");
        POITableUtil.setCellValueString(table, 0, 2, "输出功率(dBm)");
        POITableUtil.setCellValueString(table, 0, 3, "输出功率(W)");
        POITableUtil.setCellValueString(table, 0, 4, "螺流遥测(V)");
        POITableUtil.setCellValueString(table, 0, 5, "阳压遥测(V)");
        POITableUtil.setCellValueString(table, 0, 6, "开关机状态遥测(V)");
        POITableUtil.setCellValueString(table, 0, 7, "母线电流遥测（V）");
        POITableUtil.setCellValueString(table, 0, 8, "自动重启遥测（V）");

        // 遥测电压
        for (int i = 0; i < yaoCeDianYa.size(); i++) {// 每一行内容
            JsonObject jsonObjectYaoCe = yaoCeDianYa.get(i).getAsJsonObject();
            JsonObject jsonObjectMuXian = muXianDianLiu.get(i).getAsJsonObject();
            JsonObject jsonObjectPOut = ibOs.get(i).getAsJsonObject();

            double luoLIuYaoCe = jsonObjectYaoCe.get("LuoLIuYaoCe").getAsDouble();
            double yangYaYaoCe = jsonObjectYaoCe.get("YangYaYaoCe").getAsDouble();
            double kaiGuanjiYaoCe = jsonObjectYaoCe.get("KaiGuanjiYaoCe").getAsDouble();
            double muXianDianLiuYaoCe = jsonObjectYaoCe.get("MuXianDianLiuYaoCe").getAsDouble();
            double ziDongChongQiYaoCe = jsonObjectYaoCe.get("ZiDongChongQiYaoCe").getAsDouble();
            double curr = jsonObjectMuXian.get("Curr").getAsDouble();
            double pOut = jsonObjectPOut.get("POut").getAsDouble();// 输出功率
            double pOutW = jsonObjectPOut.get("POut_W").getAsDouble();// 输出功率

            POITableUtil.setCellValueDouble(table, i + 1, 1, curr);
            POITableUtil.setCellValueDouble(table, i + 1, 2, pOut);
            POITableUtil.setCellValueDouble(table, i + 1, 3, pOutW);
            POITableUtil.setCellValueDouble(table, i + 1, 4, luoLIuYaoCe);
            POITableUtil.setCellValueDouble(table, i + 1, 5, yangYaYaoCe);
            POITableUtil.setCellValueDouble(table, i + 1, 6, kaiGuanjiYaoCe);
            POITableUtil.setCellValueDouble(table, i + 1, 7, muXianDianLiuYaoCe);
            POITableUtil.setCellValueDouble(table, i + 1, 8, ziDongChongQiYaoCe);

            if (funcType == 120120) {
                JsonObject jsonObjectBoostBuck = boostBuckYaoCe.get(i).getAsJsonObject();
                JsonObject jsonObjectDeltH = deltHs.get(i).getAsJsonObject();
                JsonObject jsonObjectVoltA = voltAs.get(i).getAsJsonObject();

                double boostBuckYaoCe1 = jsonObjectBoostBuck.get("BoostBuckYaoCe").getAsDouble();
                double deltHVolt = jsonObjectDeltH.get("DeltHVolt").getAsDouble();
                double voltA = jsonObjectVoltA.get("VoltA").getAsDouble();// 输出功率

                POITableUtil.setCellValueDouble(table, i + 1, 9, boostBuckYaoCe1);
                POITableUtil.setCellValueDouble(table, i + 1, 10, deltHVolt);
                POITableUtil.setCellValueDouble(table, i + 1, 11, voltA);
            }

            POITableUtil.setCellValueDouble(table, i + 1, 0, -20 + i, 0);
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 谐波测试 = 120130
    public static void genFunc120130(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 4, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{1000, 1000, 1000});
        POITableUtil.setCellValueString(table, 0, 0, "谐波次数");
        POITableUtil.setCellValueString(table, 0, 1, "2");
        POITableUtil.setCellValueString(table, 0, 2, "3");
//修改了
        POITableUtil.setCellValueString(table, 1, 0, "指标(dBc)");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestInputData().get("xbYiZhiDu").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, data.getTestInputData().get("xbYiZhiDu").getAsDouble());

        POITableUtil.setCellValueString(table, 2, 0, "测试结果频率（GHz）");
        POITableUtil.setCellValueString(table, 3, 0, "测试结果（dBc）");
        double freq2 = data.getTestExportData().get("XB_2_Freq").getAsDouble();
        double freq3 = data.getTestExportData().get("XB_3_Freq").getAsDouble();
        double xiebo2 = data.getTestExportData().get("XB_2_Value").getAsDouble();
        double xiebo3 = data.getTestExportData().get("XB_3_Value").getAsDouble();

        if (xiebo2 > 0) {
            POITableUtil.setCellValueDouble(table, 2, 1, freq2);
            POITableUtil.setCellValueDouble(table, 3, 1, xiebo2);
        } else {
            POITableUtil.setCellValueString(table, 2, 1, "无谐波");
            POITableUtil.setCellValueString(table, 3, 1, "");
        }

        if (xiebo3 > 0) {
            POITableUtil.setCellValueDouble(table, 2, 1, freq3);
            POITableUtil.setCellValueDouble(table, 3, 1, xiebo3);
        } else {
            POITableUtil.setCellValueString(table, 2, 2, "无谐波");
            POITableUtil.setCellValueString(table, 3, 2, "");
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 带外杂波 = 120150 120160
    public static void genFunc120150(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        // zbs
        JsonArray zbs = data.getTestExportData().get("zbs").getAsJsonArray();
        int size = zbs.size();
        if (size > 3) {
            size = 3;
        }
        if (size == 0) {
            XWPFParagraph p = POITableUtil.createCursorParagraph(doc);
            p.createRun().setText("无杂波");

        } else {
            XWPFTable table = POITableUtil.createCursorTable(doc, 3, 4);
            POITableUtil.setTableColmnWidth(table, new int[]{1500, 1500, 1500, 1500});

            POITableUtil.setCellValueString(table, 0, 0, "杂波频率");
            POITableUtil.setCellValueString(table, 1, 0, "指标(dBc)");
            POITableUtil.setCellValueString(table, 2, 0, "测试结果");

            for (int i = 0; i < size; i++) {// 每一行内容
                JsonObject jsonObject = zbs.get(i).getAsJsonObject();
                double freq = jsonObject.get("freq").getAsDouble();
                POITableUtil.setCellValueDouble(table, 0, i + 1, freq);

                double level = jsonObject.get("level").getAsDouble();
                POITableUtil.setCellValueDouble(table, 2, i + 1, level);

                POITableUtil.setCellValueDouble(table, 1, i + 1, data.getTestInputData().get("zhiBiao").getAsDouble());

            }
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 矢网推饱和 = 120310 暂无数据
    public static void genFunc120310(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "频率（GHz）");
        POITableUtil.setCellValueDouble(table, 0, 1, data.getTestExportData().get("PSat_X").getAsDouble());

        POITableUtil.setCellValueString(table, 1, 0, "饱和输入功率（dBm）");
        POITableUtil.setCellValueDouble(table, 1, 1, data.getTestExportData().get("PSat_Y").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 增益平坦度 = 120920
    public static void genFunc120920(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }

        POITableUtil.createCursorParagraph(doc).createRun().setText("饱和增益平坦度及斜率");
        XWPFTable table1 = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table1, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table1, 0, 0, "项目");
        POITableUtil.setCellValueString(table1, 0, 1, "饱和增益平坦度");
        POITableUtil.setCellValueString(table1, 0, 2, "饱和增益平坦度斜率");

        POITableUtil.setCellValueString(table1, 1, 0, "指标（dB）");
        POITableUtil.setCellValueDouble(table1, 1, 1,
                data.getTestInputData().get("BaoHeZengYiPingTanDuYaoQiuZhi").getAsDouble());
        POITableUtil.setCellValueDouble(table1, 1, 2,
                data.getTestInputData().get("BaoHeZengYiXieLvYaoQiuZhi").getAsDouble());

        POITableUtil.setCellValueString(table1, 2, 0, "测试值（dB）");
        POITableUtil.setCellValueDouble(table1, 2, 1,
                data.getTestExportData().get("ZYPTD_Sat").getAsJsonObject().get("BoDong").getAsDouble());
        POITableUtil.setCellValueDouble(table1, 2, 2,
                data.getTestExportData().get("ZYPTD_Sat").getAsJsonObject().get("XieLv").getAsDouble());

        POITableUtil.createCursorParagraph(doc).createRun().setText("小信号增益平坦度及斜率");
        XWPFTable table2 = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table2, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table2, 0, 0, "项目");
        POITableUtil.setCellValueString(table2, 0, 1, "小信号增益平坦度");
        POITableUtil.setCellValueString(table2, 0, 2, "小信号增益平坦度斜率");

        POITableUtil.setCellValueString(table2, 1, 0, "指标（dB）");
        POITableUtil.setCellValueDouble(table2, 1, 1,
                data.getTestInputData().get("XiaoXinHaoZengYiPingTanDuYaoQiuZhi").getAsDouble());
        POITableUtil.setCellValueDouble(table2, 1, 2,
                data.getTestInputData().get("XiaoXinHaoZengYiXieLvYaoQiuZhi").getAsDouble());

        POITableUtil.setCellValueString(table2, 2, 0, "测试值（dB）");
        POITableUtil.setCellValueDouble(table2, 2, 1,
                data.getTestExportData().get("ZYPTD_Ss").getAsJsonObject().get("BoDong").getAsDouble());
        POITableUtil.setCellValueDouble(table2, 2, 2,
                data.getTestExportData().get("ZYPTD_Ss").getAsJsonObject().get("XieLv").getAsDouble());

        POITableUtil.createCursorParagraph(doc).createRun().setText("群时延波动及斜率");
        XWPFTable table3 = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table3, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table3, 0, 0, "项目");
        POITableUtil.setCellValueString(table3, 0, 1, "群时延波动");
        POITableUtil.setCellValueString(table3, 0, 2, "群时延斜率");

        POITableUtil.setCellValueString(table3, 1, 0, "指标（ns）");
        POITableUtil.setCellValueDouble(table3, 1, 1,
                data.getTestInputData().get("QunShiYanPingTanDuYaoQiuZhi").getAsDouble());
        POITableUtil.setCellValueDouble(table3, 1, 2,
                data.getTestInputData().get("QunShiYanXieLvYaoQiuZhi").getAsDouble());

        POITableUtil.setCellValueString(table3, 2, 0, "测试值（ns）");
        POITableUtil.setCellValueDouble(table3, 2, 1,
                data.getTestExportData().get("QSY_Sat").getAsJsonObject().get("BoDong").getAsDouble() * 1000000000D, 2);
        POITableUtil.setCellValueDouble(table3, 2, 2,
                data.getTestExportData().get("QSY_Sat").getAsJsonObject().get("XieLv").getAsDouble() * 1000000000D, 3);

        POITableUtil.createCursorParagraph(doc);
    }

    // 群时延跳变 = 120930
    public static void genFunc120930(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }

        JsonArray rowsPIn = data.getTestExportData().get("PianChaPIn").getAsJsonArray();
        JsonArray rowsPinLv = data.getTestExportData().get("PianChaPinLv").getAsJsonArray();
        JsonArray rowsPianCha = data.getTestExportData().get("PianChaPianCha").getAsJsonArray();

        int rowCnt = rowsPIn.size();

        if (data.getTestExportData().get("IsChaoCha").getAsBoolean()) {

            XWPFTable table = POITableUtil.createCursorTable(doc, rowCnt + 1, 3);
            POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
            POITableUtil.setCellValueString(table, 0, 0, "测试频率(GHz)");
            POITableUtil.setCellValueString(table, 0, 1, "测试功率(dBm)");
            POITableUtil.setCellValueString(table, 0, 2, "跳变幅值(ns)");

            for (int i = 0; i < rowCnt; i++) {
                POITableUtil.setCellValueDouble(table, i + 1, 0, rowsPinLv.get(i).getAsDouble() / 1000000000d, 3);
                POITableUtil.setCellValueDouble(table, i + 1, 1, rowsPIn.get(i).getAsDouble(), 3);
                POITableUtil.setCellValueDouble(table, i + 1, 2, rowsPianCha.get(i).getAsDouble(), 3);
            }

        } else {
            POITableUtil.createCursorParagraph(doc).createRun().setText("无超差");
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 相移_AMPM = 120940
    public static void genFunc120940(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        JsonArray xyCallBack = data.getTestExportData().get("phaseCallBackData").getAsJsonArray();
        JsonArray xyZhiBiao = data.getTestInputData().get("XiangYiHuiTuiDianZhiBiao").getAsJsonArray();
        JsonArray xyHuiTuiDian = data.getTestInputData().get("XiangYiHuiTuiDian").getAsJsonArray();

        JsonArray ampmCallBack = data.getTestExportData().get("ampmCallBackData").getAsJsonArray();
        JsonArray ampmZhiBiao = data.getTestInputData().get("AMPMZhiBiao").getAsJsonArray();
        int size = xyCallBack.size();

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 6);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 1500, 1500, 1500, 1500, 1500});
        POITableUtil.setCellValueString(table, 0, 0, "测试点");
        POITableUtil.setCellValueString(table, 1, 0, "指标（DEG）");
        POITableUtil.setCellValueString(table, 2, 0, "测试结果（DEG）");

        if (size > 5) {
            size = 5;
        }
        for (int i = 0; i < size; i++) {
            POITableUtil.setCellValueDouble(table, 0, i + 1, xyHuiTuiDian.get(i).getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, i + 1, xyZhiBiao.get(i).getAsDouble());
            POITableUtil.setCellValueDouble(table, 2, i + 1, xyCallBack.get(i).getAsDouble());
        }

        POITableUtil.createCursorParagraph(doc);

        XWPFTable table2 = POITableUtil.createCursorTable(doc, 3, 6);
        POITableUtil.setTableColmnWidth(table2, new int[]{2000, 1500, 1500, 1500, 1500, 1500});
        POITableUtil.setCellValueString(table2, 0, 0, "测试点");
        POITableUtil.setCellValueString(table2, 1, 0, "指标（DEG）");
        POITableUtil.setCellValueString(table2, 2, 0, "测试结果（DEG）");

        for (int i = 0; i < size; i++) {
            POITableUtil.setCellValueDouble(table2, 0, i + 1, xyHuiTuiDian.get(i).getAsDouble());
            POITableUtil.setCellValueDouble(table2, 1, i + 1, ampmZhiBiao.get(i).getAsDouble());
            POITableUtil.setCellValueDouble(table2, 2, i + 1, ampmCallBack.get(i).getAsDouble());
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 三阶交调 = 120950
    public static void genFunc120950(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }

        JsonArray xyCallBack = data.getTestExportData().get("tios").getAsJsonArray();
        JsonArray xyZhiBiao = data.getTestInputData().get("tioZhiBiao").getAsJsonArray();
        JsonArray xyHuiTuiDian = data.getTestInputData().get("tioCallBack").getAsJsonArray();

        int size = xyCallBack.size();

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 6);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 1500, 1500, 1500, 1500, 1500});
        POITableUtil.setCellValueString(table, 0, 0, "测试点");
        POITableUtil.setCellValueString(table, 1, 0, "指标（dBc）");
        POITableUtil.setCellValueString(table, 2, 0, "测试结果（dBc）");

        if (size > 5) {
            size = 5;
        }
        for (int i = 0; i < size; i++) {
            POITableUtil.setCellValueDouble(table, 0, i + 1, xyHuiTuiDian.get(i).getAsDouble());
            POITableUtil.setCellValueDouble(table, 1, i + 1, xyZhiBiao.get(i).getAsDouble());
            POITableUtil.setCellValueDouble(table, 2, i + 1, xyCallBack.get(i).getAsDouble());
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 带外增益 = 120960
    public static void genFunc120960(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 4);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 1000, 1000, 1000});
        POITableUtil.setCellValueString(table, 0, 0, "带外频率");
        POITableUtil.setCellValueDouble(table, 0, 1, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueDouble(table, 0, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueDouble(table, 0, 3, data.getTestExportData().get("XB_3_Value").getAsDouble());

        POITableUtil.setCellValueString(table, 1, 0, "指标（dB）");
        POITableUtil.setCellValueString(table, 1, 1, "接口单");
        POITableUtil.setCellValueString(table, 1, 2, "接口单");
        POITableUtil.setCellValueString(table, 1, 3, "接口单");

        POITableUtil.setCellValueString(table, 2, 0, "测试结果（dB）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("XB_3_Value").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 3, data.getTestExportData().get("XB_3_Value").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // ************************************************************

    // ************************************************************

    // 阳极测试结果输出 = 300020 缺少标称值、测试值、实际值
    public static void genFunc300020(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 4);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "标称值（V）");
        POITableUtil.setCellValueString(table, 0, 1, "测试值（V）");
        POITableUtil.setCellValueString(table, 0, 2, "实际值（V）");
        POITableUtil.setCellValueString(table, 0, 3, "差值（V）");

        JsonObject inData = data.getTestInputData();
        JsonObject exData = data.getTestExportData();
        POITableUtil.setCellValueDouble(table, 1, 0, inData.get("setVoltA").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 1, exData.get("actVoltA").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 2, exData.get("caleVoltA").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 3, exData.get("diffV").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 人工干预 200110
    public static void genFunc200110(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        JsonObject exData = data.getTestExportData();
        String note = "操作记录：" + exData.get("RenGongGanYu").getAsString();

        POITableUtil.createCursorParagraph(doc).createRun().setText(note);

        POITableUtil.createCursorParagraph(doc);

    }

    // 判断BoostBuck电压 500010
    public static void genFunc500010(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 4);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "调试后Boost-Buck电阻");
        POITableUtil.setCellValueString(table, 0, 1, "调试后Boost-Buck电压");
        POITableUtil.setCellValueString(table, 0, 2, "Boost-Buck电压要求值");

        JsonObject exData = data.getTestExportData();
        POITableUtil.setCellValueDouble(table, 1, 0, exData.get("BoostBuckTiaoShiQianDianZu").getAsDouble());
        POITableUtil.setCellValueDouble(table, 1, 1, exData.get("BoostBuckTiaoShiQianDianYa").getAsDouble());
        POITableUtil.setCellValueString(table, 1, 2, exData.get("TiaoShiYaoQiu").getAsString());
        POITableUtil.setCellValueString(table, 1, 3, exData.get("TiaoShiJieGuo").getAsString());

        POITableUtil.createCursorParagraph(doc);

        // 需要绘图
        // String base64 = data.getTestExportData().get("imageBase64").getAsString();
        // XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        // ImgUtil.insertForPara(para, base64, 400, 400);
        POITableUtil.createCursorParagraph(doc);

    }

    // 低压开灯丝正常状态电压_电流示波器读取 = 102020,//20190328-1650
    public static void
    genFunc102020(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }

        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);
        POITableUtil.createCursorParagraph(doc);

        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "灯丝电压有效值（V）");
        POITableUtil.setCellValueString(table, 0, 2, "灯丝电流有效值（A）");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值");
        POITableUtil.setCellValueString(table, 1, 1, "无");
        POITableUtil.setCellValueString(table, 1, 2, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("fzCurrMin").getAsDouble()) + "~"
                + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("fzCurrMax").getAsDouble()));

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("VoltRMS").getAsDouble(), 3);
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("CurrRMS").getAsDouble(), 3);
        POITableUtil.createCursorParagraph(doc);
    }

    // 关机遥测阻抗 = 102060,//20190328-1650
    public static void genFunc102060(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 6);
        POITableUtil.setTableColmnWidth(table, new int[]{1000, 1000, 1000, 1000, 1000, 1000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "螺流遥测");
        POITableUtil.setCellValueString(table, 0, 2, "阳压遥测");
        POITableUtil.setCellValueString(table, 0, 3, "开关机状态遥测");
        POITableUtil.setCellValueString(table, 0, 4, "自动重启遥测");
        POITableUtil.setCellValueString(table, 0, 5, "母线电流遥测");

        JsonObject inData = data.getTestInputData();
        JsonObject exData = data.getTestExportData();
        POITableUtil.setCellValueString(table, 1, 0, "指标要求值“≤”(KΩ)");
        POITableUtil.setCellValueString(table, 1, 1, inData.get("luoLiuYaoCeJingTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 2, inData.get("yanYaYaoJingTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 3, inData.get("kaiGuanJiYaoCeJingTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 4, inData.get("ziDongChongQiYaoCeJingTaiRes").getAsString());
        POITableUtil.setCellValueString(table, 1, 5, inData.get("muXianYaoCeJingTaiRes").getAsString());

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值(KΩ)");
        POITableUtil.setCellValueDouble(table, 2, 1, exData.get("LuoLIuYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 2, exData.get("YangYaYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 3, exData.get("KaiGuanjiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 4, exData.get("ZiDongChongQiYaoCe").getAsDouble());
        POITableUtil.setCellValueDouble(table, 2, 5, exData.get("MuXianDianLiuYaoCe").getAsDouble());

        POITableUtil.createCursorParagraph(doc);
    }

    // 提示并选择阳压截止状态高压负载阻值变化 = 500070
    // 缺少TestExportData里的数据
    public static void genFunc500070(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        // POITableUtil.createCursorParagraph(doc).createRun().setText(value);
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "要求值");
        POITableUtil.setCellValueString(table, 0, 1, "测试值");
        POITableUtil.setCellValueString(table, 0, 2, "状态");

        String zt = data.getTestExportData().get("zhuangtai").getAsString();
        double csz = data.getTestExportData().get("ceshizhi").getAsDouble();
        String yqz = data.getTestExportData().get("yaoqiuzhi").getAsString();

        POITableUtil.setCellValueString(table, 1, 0, yqz);
        POITableUtil.setCellValueDouble(table, 1, 1, csz);
        POITableUtil.setCellValueString(table, 1, 2, zt);

        // POITableUtil.setCellValueString(table, 1, 0,
        // data.getTestInputData().get("yangyaJieZhiMin").getAsString() + "~"
        // + data.getTestInputData().get("yangyaJieZhiMax").getAsString());
        // POITableUtil.setCellValueString(table, 1, 1,
        // data.getTestInputData().get("yangyaDaoTongMin").getAsString() + "~"
        // + data.getTestInputData().get("yangyaDaoTongMax").getAsString());

        // 缺少TestExportData里的数据
        // POITableUtil.setCellValueDouble(table, 2, 0,
        // data.getTestInputData().get("yangyaJieZhiMin").getAsDouble());
        // POITableUtil.setCellValueDouble(table, 2, 1,
        // data.getTestInputData().get("yangyaJieZhiMin").getAsDouble());
        POITableUtil.createCursorParagraph(doc);
    }

    // 低压欠压环路稳定性测试 = 100200,
    // 低压过压环路稳定性测试 = 100210,
    public static void genFunc100210(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 4);
        POITableUtil.setTableColmnWidth(table, new int[]{1500, 1500, 1500, 1500});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "指标要求值（V）");
        POITableUtil.setCellValueString(table, 0, 2, "实际测试值（V）");
        POITableUtil.setCellValueString(table, 0, 3, "稳定性判断");
        switch (data.getFuncType()) {
            case 100200:
                POITableUtil.setCellValueString(table, 1, 0, "低压欠压环路稳定性测试");
                POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("ovpdyqyMin").getAsDouble()) + "~"
                        + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("ovpdyqyMax").getAsDouble()));
                POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("ovp").getAsDouble());
                POITableUtil.setCellValueString(table, 1, 3, data.getTestExportData().get("resultCaption").getAsString());
                break;
            case 100210:
                POITableUtil.setCellValueString(table, 1, 0, "低压过压环路稳定性测试");
                POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("ovpdygyMin").getAsDouble()) + "~"
                        + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("ovpdygyMax").getAsDouble()));
                POITableUtil.setCellValueDouble(table, 1, 2, data.getTestExportData().get("ovp").getAsDouble());
                POITableUtil.setCellValueString(table, 1, 3, data.getTestExportData().get("resultCaption").getAsString());
        }

        POITableUtil.createCursorParagraph(doc);
    }

    // 环路稳定性测试 = 100220 高压
    public static void genFunc100220(XWPFParagraph doc, TestResultData data) {
        if (doc == null) {
            return;
        }
        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 2);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "高压环路稳定性测试");

        POITableUtil.setCellValueString(table, 1, 0, "稳定性判断");
        POITableUtil.setCellValueString(table, 1, 1, "环路正常");

        POITableUtil.createCursorParagraph(doc);
    }

    // 灯丝浪涌测试示波器读取 = 101020
    public static void genFunc101020(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);
        POITableUtil.createCursorParagraph(doc);

        XWPFTable table = POITableUtil.createCursorTable(doc, 2, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "灯丝电流有效值");
        POITableUtil.setCellValueString(table, 0, 2, "灯丝电流最大值");

        POITableUtil.setCellValueDouble(table, 1, 1,
                data.getTestExportData().get("WaveRMS_DengSiLangYong_XiuZheng").getAsDouble(), 3);
        POITableUtil.setCellValueDouble(table, 1, 2,
                data.getTestExportData().get("WaveMax_DengSiLangYong").getAsDouble(), 3);

        POITableUtil.createCursorParagraph(doc);
    }

    // 行波管灯丝稳定时间测试 = 120010, 没有

    // 行波管灯丝电流正常负载测试示波器读取 = 120030,
    public static void genFunc120030(XWPFParagraph doc, TestResultData data) throws Exception {
        if (doc == null) {
            return;
        }
        String base64 = data.getTestExportData().get("imageBase64").getAsString();
        XWPFParagraph para = POITableUtil.createCursorParagraph(doc);
        ImgUtil.insertForPara(para, base64, 400, 300);
        POITableUtil.createCursorParagraph(doc);
//修改了
        XWPFTable table = POITableUtil.createCursorTable(doc, 3, 3);
        POITableUtil.setTableColmnWidth(table, new int[]{2000, 2000, 2000});
        POITableUtil.setCellValueString(table, 0, 0, "");
        POITableUtil.setCellValueString(table, 0, 1, "灯丝电流有效值测试");
        POITableUtil.setCellValueString(table, 0, 2, "灯丝电流最大值测试");

        POITableUtil.setCellValueString(table, 1, 0, "指标要求值（A）");
        POITableUtil.setCellValueString(table, 1, 1, DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("fCurrMin").getAsDouble()) + "~"
                + DoubleUtil.DoubleZhuanHuan(data.getTestInputData().get("fCurrMax").getAsDouble()));
        // POITableUtil.setCellValueString(table, 2, 0, "指标要求值（A）");

        POITableUtil.setCellValueString(table, 2, 0, "实际测试值（A）");
        POITableUtil.setCellValueDouble(table, 2, 1, data.getTestExportData().get("CurrRMS").getAsDouble(), 3);
        POITableUtil.setCellValueDouble(table, 2, 2, data.getTestExportData().get("CurrMax").getAsDouble(), 3);

        POITableUtil.createCursorParagraph(doc);
    }
    // 阳压导通电阻设置 = 100810,

}
