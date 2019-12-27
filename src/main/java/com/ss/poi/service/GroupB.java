package com.ss.poi.service;

import com.ss.poi.entity.SensorPOIException;
import com.ss.poi.entity.TestResultData;
import com.ss.poi.util.POITableUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.ArrayList;
import java.util.List;

public class GroupB {

	public static boolean genReportZiXiangHead(XWPFDocument doc,
			TestResultData data) {
		XWPFRun run = doc.createParagraph().createRun();
		run.setText(data.getDataIndex());
		return true;
	}

	public static XWPFParagraph getParagraphToWrite(XWPFDocument doc,
			TestResultData data) throws SensorPOIException {

		String match = getMatcherByData(data);
		System.out.println("match："+match);
		XWPFParagraph para = POITableUtil.findOneParagraphByMatch(doc, match);

		if (para != null) {
			try{
				List<String> ls = getJiShu(para);
				System.out.println("DocxIndex:"+ls);
				int jishu = Integer.valueOf(ls.get(1));
				String bianhao = ls.get(2);
				String biaoti = ls.get(3);
				jishuJia1(para, jishu, jishu + 1);

				doc.insertNewParagraph(para.getCTP().newCursor()).createRun()
						.setText(bianhao + "." + jishu + " " + biaoti);
			} catch (Exception e) {
				e.printStackTrace();
				System.out.println("替换关键词出错："+match);
			}
			return doc.insertNewParagraph(para.getCTP().newCursor());
		}
		throw new SensorPOIException(-1);
	}

	private static List<String> getJiShu(XWPFParagraph para) {
		int num2, num3;
		String data = para.getText();
		ArrayList<String> ls = new ArrayList<>();

		while (data.contains("[")) {
			num2 = data.indexOf("[");
			num3 = data.indexOf("]");
			data = data.replaceFirst("\\[", "*");
			data = data.replaceFirst("\\]", "*");
			String s = data.substring(num2 + 1, num3);
			ls.add(s);
		}

		return ls;
		// return
		// Integer.valueOf(para.getText().substring(para.getText().indexOf("[")+1,para.getText().indexOf("]")));
	}
	
	public static void replaceDateTime(XWPFDocument doc, String dateTime) {
		String oldStr = "{{ShiJian}}";
		replaceOneString(doc,oldStr,dateTime);
	}
	
	public static void replaceUserName(XWPFDocument doc, String userName) {
		String oldStr = "{{CeShiRenYuan}}";
		replaceOneString(doc,oldStr,userName);
	}

	public static void replaceXingHao(XWPFDocument doc, String xingHao) {
		String oldStr = "{{XingHao}}";
		replaceOneString(doc,oldStr,xingHao);
	}

	public static void replaceJieDuan(XWPFDocument doc, String jieDuan) {
		String oldStr = "{{JieDuan}}";
		replaceOneString(doc,oldStr,jieDuan);
	}

	public static void replaceTouChanBianHao(XWPFDocument doc, String touChanBianHao) {
		String oldStr = "{{TouChanBianHao}}";
		replaceOneString(doc,oldStr,touChanBianHao);
	}
	
	private static void replaceOneString(XWPFDocument doc, String oldStr,String newStr){

		List<XWPFParagraph> paras = POITableUtil.findAllParagraphByMatch(doc,
				oldStr);
		for (int i = 0; i < paras.size(); i++) {
			replaceInParagraph(paras.get(i), oldStr, newStr);
		}
	}

	public static void replaceUnUsed(XWPFDocument doc) {
		String oldStr = "{{";

		List<XWPFParagraph> paras = POITableUtil.findAllParagraphByMatch(doc,
				oldStr);
		for (int i = 0; i < paras.size(); i++) {
			removeParaRuns(paras.get(i));
		}
	}

	public static void ClearUnusedIndex(XWPFDocument doc) {

	}

	private static void replaceInParagraph(XWPFParagraph para, String oldStr,
			String newStr) {
		String newIndex = para.getText().replace(oldStr, newStr);
		removeParaRuns(para);
		para.createRun().setText(newIndex);
	}
	
	private static void removeParaRuns(XWPFParagraph para){

		//for (int i = 0; i < para.getRuns().size(); i++) {
		//	para.removeRun(0);
		//}
		for (int i = (para.getRuns().size()-1); i >=0 ; i--) {
			para.removeRun(i);
		}
	}

	private static void jishuJia1(XWPFParagraph para, int oldJiShu, int newJiShu) {
		String newIndex = para.getText().replace("[" + oldJiShu + "]",
				"[" + newJiShu + "]");
		System.out.println("jishuJia1_para.getRuns().size():"+para.getRuns().size());

		removeParaRuns(para);
		/*for (int i = 0; i < para.getRuns().size(); i++) {
			para.removeRun(0);
		}		for (int i = (para.getRuns().size()-1); i >=0 ; i--) {
			para.removeRun(i);
		}*/
		para.createRun().setText(newIndex);
	}

	public static String getMatcherByData(TestResultData data) {
		return "["+data.getDataIndex2().substring(
				data.getDataIndex2().indexOf("_") + 1)+"]";
	}

	public static void main(String[] args) {
		String data = "0_高压初测流程_低压调试框-灯丝正常电流测试（负载）_低压开灯丝电流正常负载测试示波器读取[13][1.2.5]";
		int num = 0;
		int num2 = 0;
		int num3 = 0;
		ArrayList<Integer> list = new ArrayList<>();
		ArrayList<Integer> list2 = new ArrayList<>();
		while (data.contains("[")) {
			num2 = data.indexOf("[");
			num3 = data.indexOf("]");
			data = data.replaceFirst("\\[", "*");
			data = data.replaceFirst("\\]", "*");
			System.out.println("data" + data);
			list.add(num2);
			list2.add(num3);
			System.out.println("num2:" + num2);
			System.out.println("num3:" + num3);
			num++;
		}
		System.out.println(num);
		for (int i = 0; i < num; i++) {
			String number = data.substring(list.get(i) + 1, list2.get(i));
			System.out.println("number:" + number);
		}
		Integer of = Integer.valueOf(data.substring(data.indexOf("[") + 1,
				data.indexOf("]")));
		System.out.println(of);

	}
}
