package com.ss.poi.test;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.ss.poi.util.ReportFileUtil;

/**
 * @Description
 */
public class WordTextTest {
	public static void main(String[] args) throws Exception {
		XWPFDocument doc = new XWPFDocument();

		XWPFParagraph para;
		XWPFRun run;
		// 添加文本
		String content = "额尔古纳河在1689年的《中俄尼布楚条约》中成为中国和俄罗斯的界河，额尔古纳河上游称海拉尔河，源于大兴安岭西侧，西流至阿该巴图山脚， 折而北行始称额尔古纳河。额尔古纳河在黑龙江省漠河县以西的内蒙古自治区额尔古纳右旗的恩和哈达附近与流经俄罗斯境内的石勒喀河汇合后始称黑龙江。沿额尔古纳河沿岸地区土地肥沃，森林茂密，水草丰美， 鱼类品种很多，动植物资源丰富，宜农宜木，是人类理想的天堂。";
		para = doc.createParagraph();
		para.setAlignment(ParagraphAlignment.CENTER);// 设置左对齐 
		para.setIndentationFirstLine(1000);
		run = para.createRun();
		run.setFontFamily("仿宋");
		run.setFontSize(13);
		run.setText(content);
		doc.createParagraph(); // 添加图片
		String[] imgs = { "D:\\asd.png", "D:\\asd.png" };
		for (int i = 0; i < imgs.length; i++) {
			para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);// 设置左对齐 
			run =para.createRun();
			InputStream input = new FileInputStream(imgs[i]);
			run.addPicture(input, XWPFDocument.PICTURE_TYPE_JPEG, imgs[i],
					Units.toEMU(35), Units.toEMU(17));
			para = doc.createParagraph();
			para.setAlignment(ParagraphAlignment.CENTER);// 设置左对齐 
			run =para.createRun();
			run.setFontFamily("仿宋");
			run.setFontSize(11);
			run.setText(imgs[i]);
		}
		doc.createParagraph();
		ReportFileUtil.saveFile(doc, "123tyu", "DOC");
	}

}
