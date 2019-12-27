package com.ss.poi.test;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import com.ss.poi.util.ImgUtil;
import com.ss.poi.util.ReportFileUtil;

/**
 * @Description 
 */
public class InsertImgTest {
	public static void main(String[] args) throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFTable table = doc.createTable(3, 3);
		XWPFParagraph paraImg =table.getRow(0).getCell(0).getParagraphs().get(0);//对某个单元格设置段落，
		String base64 = ReportFileUtil.encodeBase64File("e:/asd.png");
		ImgUtil.insertForPara(paraImg, base64,400, 400);
		doc.write(new FileOutputStream("e:/a.docx"));
		doc.close();
	}

}
