package com.ss.poi.controller;

import com.ss.poi.entity.TestResultData;
import com.ss.poi.service.GroupB;
import com.ss.poi.service.GroupCA;
import com.ss.poi.util.GsonUtil;
import com.ss.poi.util.ReportFileUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * @Description
 */
public class WordReport {

	private static XWPFDocument checkAndOpenDocx(String fileLocation, TestResultData data) throws Exception {
		if (!ReportFileUtil.checkFileLocationExists(fileLocation)) {
			String tplLocation = ReportFileUtil.getLiuChengTempleteLocation(data);
			ReportFileUtil.copyFile(tplLocation, fileLocation);
			XWPFDocument doc = ReportFileUtil.openDocx(fileLocation);

			GroupB.replaceXingHao(doc, data.getTestXingHao());
			GroupB.replaceJieDuan(doc, data.getTestJieDuan());
			GroupB.replaceTouChanBianHao(doc, data.getTestedID());
			GroupB.replaceDateTime(doc, data.getDateStr());
			GroupB.replaceUserName(doc, data.getUserName());

			return doc;
		}
		return ReportFileUtil.openDocx(fileLocation);
	}

	public static boolean creatWordReport(TestResultData data) throws Exception {
		String fileLocation_History = ReportFileUtil.getFileLocationByData(data);
		String fileLocation_Process = ReportFileUtil.getFileLocationByMissionID(data.getMissionID());

		XWPFDocument docZiXiang = ReportFileUtil.openDocx(fileLocation_History);
		XWPFDocument docProcess = checkAndOpenDocx(fileLocation_Process, data);// ReportFileUtil.openDocx(fileLocation_Process);

		GroupB.genReportZiXiangHead(docZiXiang, data);
		// GroupB.genReportZiXiangHead(docProcess, data);
		System.out.println("123:" + data.funcType);

		switch (data.funcType) {
		case 100030:
			GroupCA.genFunc100030(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100030(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100040:
			GroupCA.genFunc100040(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100040(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100050:
			GroupCA.genFunc100050(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100050(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100060:
			GroupCA.genFunc100060(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100060(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100070:
			GroupCA.genFunc100070(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100070(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100080:
			GroupCA.genFunc100080(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100080(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100130:
		case 300030:
			GroupCA.genFunc100130(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100130(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100150:
			GroupCA.genFunc100150(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100150(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100160:
		case 100190:
			GroupCA.genFunc100190(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100190(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100180:
			GroupCA.genFunc100180(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100180(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100200:
		case 100210:
			GroupCA.genFunc100210(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100210(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100220:
			GroupCA.genFunc100220(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100220(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100230:
			GroupCA.genFunc100230(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100230(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100240:
			GroupCA.genFunc100240(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100240(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100260:
			GroupCA.genFunc100260(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100260(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100270:
			GroupCA.genFunc100270(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100270(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100280:
			GroupCA.genFunc100280(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100280(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100290:
			GroupCA.genFunc100290(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100290(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 100530:
			GroupCA.genFunc100530(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100530(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100570:
			GroupCA.genFunc100570(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100570(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100610:
			GroupCA.genFunc100610(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100610(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100620:
			GroupCA.genFunc100620(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100620(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100630:
			GroupCA.genFunc100630(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100630(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100640:
		case 100680:
			GroupCA.genFunc100640(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100640(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100650:
			GroupCA.genFunc100650(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100650(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 100660:
			GroupCA.genFunc100660(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100660(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 100710:
			GroupCA.genFunc100710(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100710(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 100810:
		case 100820:
			GroupCA.genFunc100810(docZiXiang.createParagraph(), data);
			GroupCA.genFunc100810(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 102020:
		case 102040:
			GroupCA.genFunc102020(docZiXiang.createParagraph(), data);
			GroupCA.genFunc102020(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 102060:
			GroupCA.genFunc102060(docZiXiang.createParagraph(), data);
			GroupCA.genFunc102060(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 102080:
			GroupCA.genFunc102080(docZiXiang.createParagraph(), data);
			GroupCA.genFunc102080(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 102090:
			GroupCA.genFunc102090(docZiXiang.createParagraph(), data);
			GroupCA.genFunc102090(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 101020:
			GroupCA.genFunc101020(docZiXiang.createParagraph(), data);
			GroupCA.genFunc101020(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120030:
			GroupCA.genFunc120030(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120030(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120050:
			GroupCA.genFunc120050(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120050(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120110:
			GroupCA.genFunc120110(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120110(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120120:
		case 120170:
			GroupCA.genFunc120120(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120120(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120130:
			GroupCA.genFunc120130(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120130(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 120150:
		case 120160:
			GroupCA.genFunc120150(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120150(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120310:
			GroupCA.genFunc120310(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120310(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 120920:
			GroupCA.genFunc120920(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120920(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120930:
			GroupCA.genFunc120930(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120930(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120940:
			GroupCA.genFunc120940(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120940(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 120950:
			GroupCA.genFunc120950(docZiXiang.createParagraph(), data);
			GroupCA.genFunc120950(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 200110:
			GroupCA.genFunc200110(docZiXiang.createParagraph(), data);
			GroupCA.genFunc200110(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 300020:
			GroupCA.genFunc300020(docZiXiang.createParagraph(), data);
			GroupCA.genFunc300020(GroupB.getParagraphToWrite(docProcess, data), data);
			break;

		case 500070:
			GroupCA.genFunc500070(docZiXiang.createParagraph(), data);
			GroupCA.genFunc500070(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		case 500010:
			GroupCA.genFunc500010(docZiXiang.createParagraph(), data);
			GroupCA.genFunc500010(GroupB.getParagraphToWrite(docProcess, data), data);
			break;
		default:
			break;
		}

		ReportFileUtil.saveAsDocx(docZiXiang, fileLocation_History);
		ReportFileUtil.saveAsDocx(docProcess, fileLocation_Process);

		GroupB.replaceUnUsed(docProcess);
		ReportFileUtil.saveAsDocx(docProcess, fileLocation_Process + ".docx");

		return true;
	}

	public static boolean creatWordReport(String dataStr) throws Exception {

		TestResultData data = GsonUtil.jsonToObject(dataStr, TestResultData.class);
		return creatWordReport(data);
	}
}
