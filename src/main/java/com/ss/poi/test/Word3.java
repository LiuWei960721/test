package com.ss.poi.test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

/**
 * @Description 
 */
public class Word3 {
	public static void main(String[] args) throws IOException {
		XWPFDocument document = new XWPFDocument();
		XWPFTable table = document.createTable(4,4);
		XWPFParagraph para =table.getRow(1).getCell(0).addParagraph();//对某个单元格设置段落�?
		para.setAlignment(ParagraphAlignment.CENTER);
		XmlCursor newCursor = para.getCTP().newCursor();
		XWPFTable table2 = para.getBody().insertNewTbl(newCursor);
		
		//XmlCursor newCursor =document.getDocument().getBody().getTblArray(0).newCursor();
		int[] rowHeights = new int[] { 180, 1680, 157, 1100 };
		for (int i = 0; i < rowHeights.length; i++) {
			table.getRow(i).setHeight(rowHeights[i]);

		}
		/*table.getRow(0).setHeight(180);
		table.getRow(1).setHeight(1680);
		table.getRow(2).setHeight(157);
		table.getRow(3).setHeight(1100);*/
			
		CTTbl ttbl = table.getCTTbl();
		CTTblGrid tblGrid = ttbl.addNewTblGrid();
		int[] colWidths = new int[] { 4000, 1000, 1200, 1300 };
		for (int i : colWidths) {
			CTTblGridCol gridCol = tblGrid.addNewGridCol();
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
        
		CTTbl ttbl2 = table2.getCTTbl();
		CTTblGrid tblGrid2 = ttbl2.addNewTblGrid();
		int[] colWidths2 = new int[] { 1600, 1600 };
		for (int i : colWidths2) {
			CTTblGridCol gridCol2 = tblGrid2.addNewGridCol();
			gridCol2.setW(new BigInteger(i + ""));
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
        
        
   /* 	table2.getRow(0).setHeight(20);
		table2.getRow(1).setHeight(20);
		table2.getRow(2).setHeight(20);
		table2.getRow(3).setHeight(20);*/
		
		String path = "D:\\123asd.doc";
		OutputStream os = new FileOutputStream(path);
		document.write(os);
	}

}
