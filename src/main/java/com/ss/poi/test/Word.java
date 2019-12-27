package com.ss.poi.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

/**
 * @Description
 */
public class Word {
	public static void main(String[] args) throws Exception {
		XWPFDocument doc = new XWPFDocument();
		/*
		 * // 添加文本 String content =
		 * "额尔古纳河在1689年的《中俄尼布楚条约》中成为中国和俄罗斯的界河，额尔古纳河上游称海拉尔河，源于大兴安岭西侧，西流至阿该巴图山脚， 折而北行始称额尔古纳河。额尔古纳河在黑龙江省漠河县以西的内蒙古自治区额尔古纳右旗的恩和哈达附近与流经俄罗斯境内的石勒喀河汇合后始称黑龙江。沿额尔古纳河沿岸地区土地肥沃，森林茂密，水草丰美， 鱼类品种很多，动植物资源丰富，宜农宜木，是人类理想的天堂。"
		 * ; para = doc.createParagraph();
		 * para.setAlignment(ParagraphAlignment.LEFT);// 设置左对齐 run =
		 * para.createRun(); run.setFontFamily("仿宋"); run.setFontSize(13);
		 * run.setText(content); doc.createParagraph(); // 添加图片 String[] imgs =
		 * { "D:\\a.jpg", "D:\\b.jpg" }; for (int i = 0; i < imgs.length; i++) {
		 * para = doc.createParagraph();
		 * para.setAlignment(ParagraphAlignment.CENTER);// 设置左对齐 run =
		 * para.createRun(); InputStream input = new FileInputStream(imgs[i]);
		 * run.addPicture(input, XWPFDocument.PICTURE_TYPE_JPEG, imgs[i],
		 * Units.toEMU(350), Units.toEMU(170)); para = doc.createParagraph();
		 * para.setAlignment(ParagraphAlignment.CENTER);// 设置左对齐 run =
		 * para.createRun(); run.setFontFamily("仿宋"); run.setFontSize(11);
		 * run.setText(imgs[i]); } doc.createParagraph();
		 */

		// 添加表格
		XWPFTable table = doc.createTable(4, 4);
		table.setCellMargins(5, 3, 5, 3);
		// 设置指定宽度
		CTTbl ttbl = table.getCTTbl();
		CTTblGrid tblGrid = ttbl.addNewTblGrid();
		int[] colWidths = new int[] { 4140, 1545, 2595, 2415 };
		for (int i : colWidths) {
			CTTblGridCol gridCol = tblGrid.addNewGridCol();
			gridCol.setW(new BigInteger(i + ""));
		}
		// 设置高度
		
		int[] rowHeights = new int[] { 800, 800, 570, 11015 };
		for (int i = 0; i < rowHeights.length; i++) {
			table.getRow(i).setHeight(rowHeights[i]);

		}
		
		/*
		 * for(int i=1; i<table.getNumberOfRows(); i++){
		 * table.getRow(i).setHeight(5100); }
		 */
		mergeCellVertically(table, 3, 0, 2);
		mergeCellHorizontally(table, 1, 1, 2);
		mergeCellHorizontally(table, 2, 1, 2);
		mergeCellHorizontally(table, 3, 0, 3);

		setTableLocation(table, "center");
		// table.addNewCol();//添加新列
		// table.createRow();//添加新行
		// String[] title = new String[] { "境内河流", "境外河流", "合计", "s" };
		String[] value = new String[] { " 型号：", " 阶段：", " 投产编号：", "", " 工序：",
				" 测试项：", "日期：", "测试人员：" };
		XWPFTableRow row_0 = table.getRow(0);
		row_0.getCell(0).setText(value[0]);
		
		XWPFParagraph para1 =table.getRow(0).getCell(0).addParagraph();//对某个单元格设置段落，
		XWPFRun run1 = para1.createRun();
		run1.setText("KA频段55W");
		run1.setBold(false);
		run1.setFontSize(12);
		para1.setAlignment(ParagraphAlignment.CENTER);

		
		XWPFTableRow row_1 = table.getRow(0);
		row_1.getCell(1).setText(value[1]);
		
		XWPFTableRow row_2 = table.getRow(1);
		row_2.getCell(1).setText(value[5]);
		
//       img.setAbsolutePosition(0, 0);  
//       img.setAlignment(Image.MIDDLE);//设置图片显示位置 
		
		XWPFParagraph paraImg =table.getRow(0).getCell(3).getParagraphs().get(0);//对某个单元格设置段落，
		paraImg.createRun().addPicture(new FileInputStream("d:/asd.png"), XWPFDocument.PICTURE_TYPE_PNG, "",
				Units.toEMU(50), Units.toEMU(50));
		paraImg.setAlignment(ParagraphAlignment.CENTER);
		paraImg.setVerticalAlignment(TextAlignment.CENTER);
		
		//表格嵌套
		XWPFParagraph para =table.getRow(3).getCell(0).addParagraph();//对某个单元格设置段落，
		para.setAlignment(ParagraphAlignment.CENTER);
		XmlCursor newCursor = para.getCTP().newCursor();
		XWPFTable table2 = para.getBody().insertNewTbl(newCursor);
		
			
		CTTbl ttbl2 = table.getCTTbl();
		CTTblGrid tblGrid2 = ttbl2.addNewTblGrid();
		int[] colWidths2 = new int[] { 4000, 1000, 1200, 1300 };
		for (int i : colWidths2) {
			CTTblGridCol gridCol = tblGrid2.addNewGridCol();
			gridCol.setW(new BigInteger(i + ""));
		}
		
		/*
		CTTblWidth addNewTblW = table2.getCTTbl().addNewTblPr().addNewTblW();
		addNewTblW.setType(STTblWidth.DXA);
		addNewTblW.setW(BigInteger.valueOf(9072));*/
		
		/*CTTbl ttbl2 = table2.getCTTbl();
		CTTblPr tblPr = ttbl2.getTblPr() == null ? ttbl2.addNewTblPr() : ttbl2.getTblPr();
		CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
		CTJc cTJc=tblPr.addNewJc();
		cTJc.setVal(STJc.Enum.forString("center"));
		tblWidth.setW(new BigInteger("400"));
		tblWidth.setType(STTblWidth.DXA);*/
		
		System.out.println("table:"+table);
		System.out.println("table2:"+table2);
		XWPFTableRow row_00 = table2.createRow();
        //row_00.getCell(0).setText("姓名");
		
		
        row_00.addNewTableCell().setText("年龄");
        row_00.createCell().setText("姓名");

        row_00.setHeight(200);
        
        XWPFTableCell cell = row_00.getCell(0);

        /** 设置水平居中 */
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        
		XWPFTableRow row_01 = table2.createRow();
        row_01.getCell(0).setText("aa");
        row_01.getCell(1).setText("bb");
        row_01.setHeight(200);
        
		CTTbl ttbl3 = table2.getCTTbl();
		CTTblGrid tblGrid3 = ttbl3.addNewTblGrid();
		int[] colWidths3 = new int[] { 1600, 1600 };
		for (int i : colWidths3) {
			CTTblGridCol gridCol3 = tblGrid3.addNewGridCol();
			gridCol3.setW(new BigInteger(i + ""));
		}
		
		
		CTTblBorders borders=table2.getCTTbl().addNewTblPr().addNewTblBorders();  

		
		CTBorder hBorder=borders.addNewInsideH();  
		hBorder.setVal(STBorder.Enum.forString("single"));  
		hBorder.setSz(new BigInteger("1"));  
		hBorder.setColor("000000");   
          
        CTBorder vBorder=borders.addNewInsideV();  
        vBorder.setVal(STBorder.Enum.forString("single"));  
        vBorder.setSz(new BigInteger("1"));  
        vBorder.setColor("000000");  
		
		
        CTBorder lBorder=borders.addNewLeft();  
        lBorder.setVal(STBorder.Enum.forString("single"));  
        lBorder.setSz(new BigInteger("1"));  
        lBorder.setColor("000000");  
          
        CTBorder rBorder=borders.addNewRight();  
        rBorder.setVal(STBorder.Enum.forString("single"));  
        rBorder.setSz(new BigInteger("1"));  
        rBorder.setColor("000000");  
          
        CTBorder tBorder=borders.addNewTop();  
        tBorder.setVal(STBorder.Enum.forString("single"));  
        tBorder.setSz(new BigInteger("1"));  
        tBorder.setColor("000000");  
          
        CTBorder bBorder=borders.addNewBottom();  
        bBorder.setVal(STBorder.Enum.forString("single"));  
        bBorder.setSz(new BigInteger("1"));  
        bBorder.setColor("000000");  
		
		//设置列宽跟随内容伸缩
		/*CTTblWidth infoTableWidth = table2.getCTTbl().addNewTblPr().addNewTblW();  
        infoTableWidth.setType(STTblWidth.DXA);  
        infoTableWidth.setW(BigInteger.valueOf(9072));*/ 
		
        Word.setTableLocation(table2, "center");
		
		

        //setCellLocation(table,"CENTER","center");
        
        
		
		/*XWPFParagraph para = doc.createParagraph();
		XWPFRun run = para.createRun();
		para.setAlignment(ParagraphAlignment.CENTER);// 设置左对齐 run =
		InputStream input = new FileInputStream("D:\\asd.png");
		run.addPicture(input, XWPFDocument.PICTURE_TYPE_PNG, "",
				Units.toEMU(50), Units.toEMU(50));
		XWPFTableRow rowImg = table.getRow(0);
		rowImg.getCell(3).setParagraph(para);
		para.removeRun(0);*/
		
		/*XWPFTableCell cell = table.getRow(1).getCell(1);
		CTTc cttc = cell.getCTTc();
		IBody body = table.getBody().getTableCell(cttc);
        XmlCursor tblCursor = doc.getDocument().getBody().getTblArray(0).newCursor(); //使得游标获得了第一个表格的位置
        XmlCursor cursor = para .getCTP().newCursor();
		XWPFTable insertNewTbl = body.insertNewTbl(cursor);
		insertNewTbl.addNewCol();*/
		
		
		
		
		/*XWPFTableCell tableCell = body.getTableCell(cttc);
		tableCell.*/
		
		/*XWPFParagraph p = doc.getParagraphArray(0);
        p.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run2 = p.insertNewRun(0);
		XWPFTableCell imageCell = table.getRow(0).getCell(0);
        List<XWPFParagraph> paragraphs = imageCell.getParagraphs();
        XWPFParagraph newPara = paragraphs.get(0);
        XWPFRun imageCellRunn = newPara.createRun();
        imageCellRunn.addPicture(new FileInputStream("D:\\a.jpg"), XWPFDocument.PICTURE_TYPE_PNG, "1.png", Units.toEMU(60), Units.toEMU(30));
        run2.addBreak();*/
		
		

		String path = "e:\\test517.docx";
		OutputStream os = new FileOutputStream(path);
		doc.write(os);
		if (os != null) {
			try {
				os.close();
				System.out.println("文件已输出！");
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 合并行
	 * 
	 * @param table
	 * @param col
	 *            需要合并的列
	 * @param fromRow
	 *            开始行
	 * @param toRow
	 *            结束行
	 */
	public static void mergeCellVertically(XWPFTable table, int col,
			int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			CTVMerge vmerge = CTVMerge.Factory.newInstance();
			if (rowIndex == fromRow) {
				vmerge.setVal(STMerge.RESTART);
			} else {
				vmerge.setVal(STMerge.CONTINUE);
			}
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			CTTcPr tcPr = cell.getCTTc().getTcPr();
			if (tcPr != null) {
				tcPr.setVMerge(vmerge);
			} else {
				tcPr = CTTcPr.Factory.newInstance();
				tcPr.setVMerge(vmerge);
				cell.getCTTc().setTcPr(tcPr);
			}
		}
	}

	/**
	 * 列合并
	 * 
	 * @param table
	 * @param row
	 *            需要合并的列
	 * @param fromCol
	 *            开始行
	 * @param toCol
	 *            结束行
	 */
	public static void mergeCellHorizontally(XWPFTable table, int row,
			int fromCol, int toCol) {
		for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
			CTHMerge hmerge = CTHMerge.Factory.newInstance();
			if (colIndex == fromCol) {
				hmerge.setVal(STMerge.RESTART);
			} else {
				hmerge.setVal(STMerge.CONTINUE);
			}
			XWPFTableCell cell = table.getRow(row).getCell(colIndex);
			CTTcPr tcPr = cell.getCTTc().getTcPr();
			if (tcPr != null) {
				tcPr.setHMerge(hmerge);
			} else {
				tcPr = CTTcPr.Factory.newInstance();
				tcPr.setHMerge(hmerge);
				cell.getCTTc().setTcPr(tcPr);
			}
		}
	}

	/**
	 * 设置表格位置
	 * 
	 * @param xwpfTable
	 * @param location
	 *            整个表格居中center,left居左，right居右，both两端对齐
	 */
	public static void setTableLocation(XWPFTable xwpfTable, String location) {
		CTTbl cttbl = xwpfTable.getCTTbl();
		CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl
				.getTblPr();
		CTJc cTJc = tblpr.addNewJc();
		cTJc.setVal(STJc.Enum.forString(location));
	}

	/**
	 * 设置单元格水平位置和垂直位置
	 * 
	 * @param xwpfTable
	 * @param verticalLoction
	 *            单元格中内容垂直上TOP，下BOTTOM，居中CENTER，BOTH两端对齐
	 * @param horizontalLocation
	 *            单元格中内容水平居中center,left居左，right居右，both两端对齐
	 */
	public static void setCellLocation(XWPFTable xwpfTable,
			String verticalLoction, String horizontalLocation) {
		List<XWPFTableRow> rows = xwpfTable.getRows();
		for (XWPFTableRow row : rows) {
			List<XWPFTableCell> cells = row.getTableCells();
			for (XWPFTableCell cell : cells) {
				CTTc cttc = cell.getCTTc();
				CTP ctp = cttc.getPList().get(0);
				CTPPr ctppr = ctp.getPPr();
				if (ctppr == null) {
					ctppr = ctp.addNewPPr();
				}
				CTJc ctjc = ctppr.getJc();
				if (ctjc == null) {
					ctjc = ctppr.addNewJc();
				}
				ctjc.setVal(STJc.Enum.forString(horizontalLocation)); // 水平居中
				cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign
						.valueOf(verticalLoction));// 垂直居中
			}
		}
	}
}
