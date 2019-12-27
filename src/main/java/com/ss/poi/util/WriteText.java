package com.ss.poi.util;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * @Description
 */
public class WriteText {
	/**
	 * 在新的段落写文字,默认居中对齐
	 * 
	 * @param doc
	 * @param content
	 * @param spaceNum 首行缩进量
	 */
	public static XWPFParagraph writeForPara(XWPFDocument doc, String content,
			int spaceNum) throws Exception {
		XWPFParagraph para = doc.createParagraph();
		para.createRun().setText(content);
		para.setIndentationFirstLine(spaceNum);
		para.setAlignment(ParagraphAlignment.CENTER);
		return para;
	}
}
