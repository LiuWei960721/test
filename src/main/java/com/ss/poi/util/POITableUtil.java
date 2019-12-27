package com.ss.poi.util;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Description
 */
public class POITableUtil {

	public static XWPFParagraph findOneParagraphByMatch(XWPFDocument doc,
			String match) {
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para = null;
		while (iterator.hasNext()) {
			para = iterator.next();
			if (hasMatcherInParagraph(para, match)) {
				return para;
			}
		}

		Iterator<XWPFTable> iteratorTable = doc.getTablesIterator();
		XWPFTable tabl = null;
		while (iteratorTable.hasNext()) {
			tabl = iteratorTable.next();

			int rowCnt = tabl.getRows().size();
			for (int r = 0; r < rowCnt; r++) {
				XWPFTableRow row = tabl.getRow(r);
				int colCnt = row.getTableCells().size();
				for (int c = 0; c < colCnt; c++) {
					XWPFTableCell cell = row.getCell(c);
					int parCnt = cell.getParagraphs().size();
					for (int p = 0; p < parCnt; p++) {
						para = cell.getParagraphs().get(p);
						if (hasMatcherInParagraph(para, match)) {
							return para;
						}
					}
				}
			}
		}

		return null;
	}
	public static List<XWPFParagraph> findAllParagraphByMatch(XWPFDocument doc,
			String match) {
		List<XWPFParagraph> lp=new ArrayList<>();
		
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para = null;
		while (iterator.hasNext()) {
			para = iterator.next();
			if (hasMatcherInParagraph(para, match)) {
				lp.add(para);
			}
		}

		Iterator<XWPFTable> iteratorTable = doc.getTablesIterator();
		XWPFTable tabl = null;
		while (iteratorTable.hasNext()) {
			tabl = iteratorTable.next();

			int rowCnt = tabl.getRows().size();
			for (int r = 0; r < rowCnt; r++) {
				XWPFTableRow row = tabl.getRow(r);
				int colCnt = row.getTableCells().size();
				for (int c = 0; c < colCnt; c++) {
					XWPFTableCell cell = row.getCell(c);
					int parCnt = cell.getParagraphs().size();
					for (int p = 0; p < parCnt; p++) {
						para = cell.getParagraphs().get(p);
						if (hasMatcherInParagraph(para, match)) {
							lp.add(para);
						}
					}
				}
			}
		}

		return lp;
	}

	private static boolean hasMatcherInParagraph(XWPFParagraph para,
			String match) {
		String text = getParagraphText(para);
		if (text.contains(match)) {
			return true;
		}
		return false;
	}
	
	public static String getParagraphText(XWPFParagraph para){
		List<XWPFRun> runs= para.getRuns();
		String text="";
		for (int i = 0; i < runs.size(); i++) {
			XWPFRun run = runs.get(i);
			text = text + run.toString();
		}
		return text;
	}

	public static void replaceInPara(XWPFDocument doc,
			Map<String, Object> params) {
		Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
		XWPFParagraph para;
		while (iterator.hasNext()) {
			para = iterator.next();
			replaceInPara(para, params);
		}
	}

	public static void replaceInPara(XWPFParagraph para,
			Map<String, Object> params) {
		List<XWPFRun> runs;
		Matcher matcher;
		if (matcher(para.getParagraphText()).find()) {
			runs = para.getRuns();
			int start = -1;
			int end = -1;
			String str = "";
			for (int i = 0; i < runs.size(); i++) {
				XWPFRun run = runs.get(i);
				String runText = run.toString().trim();
				if (!"".equals(runText)) {
					if ('$' == runText.charAt(0) && '{' == runText.charAt(1)) {
						start = i;
					}
					if ((start != -1)) {
						str += runText;
					}
					if ('}' == runText.charAt(runText.length() - 1)) {
						if (start != -1) {
							end = i;
							break;
						}
					}
				}
			}
			if (start != -1) {
				for (int i = start; i < end + 1; i++) {
					para.removeRun(start);
				}
				XWPFRun createRun = para.insertNewRun(start);
				for (String key : params.keySet()) {
					if (str.equals(key)) {
						createRun.setText((String) params.get(key));
						createRun.setFontSize(16);
						createRun.addBreak();
						break;
					}
				}
			}

		}
	}

	private static Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}",
				Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}

	private static void changeMessage(Map<String, String> params,
			XWPFDocument document) {
		Iterator<XWPFTable> itTable = document.getTablesIterator();
		while (itTable.hasNext()) {
			XWPFTable table = (XWPFTable) itTable.next();
			changeTableMessage(params, table);
		}
	}

	private static void changeTableMessage(Map<String, String> params,
			XWPFTable table) {// , boolean isBold, Integer fontSize
		int count = table.getNumberOfRows();// 获取table的行数
		for (int i = 0; i < count; i++) {
			XWPFTableRow row = table.getRow(i);
			List<XWPFTableCell> cells = row.getTableCells();
			for (XWPFTableCell cell : cells) {// 遍历每行的值并进行替换
				for (Map.Entry<String, String> e : params.entrySet()) {
					if (cell.getText().equals(e.getKey())) {
						XWPFParagraph newPara = new XWPFParagraph(cell
								.getCTTc().addNewP(), cell);
						XWPFRun r1 = newPara.createRun();
						// r1.setBold(isBold);
						// if (fontSize != null) {
						// r1.setFontSize(fontSize);
						// }
						r1.setText(e.getValue());
						cell.removeParagraph(0);
						cell.setParagraph(newPara);
					}
				}
			}
		}
	}

	public static XWPFTable createCommonTable(XWPFDocument doc, int rowCnt,
			int colCnt) {

		XWPFTable table = doc.createTable(rowCnt, colCnt);
		setLocation(table, 2);
		// table.setWidth(2000);
		return table;
	}

	public static XWPFTable createCursorTable(XWPFParagraph prgh, int rowCnt,
			int colCnt) {

		XmlCursor cursor = prgh.getCTP().newCursor();
		XWPFTable tableOne = prgh.getBody().insertNewTbl(cursor);
		XWPFTableRow row = tableOne.getRow(0);
		for (int c = 0; c < colCnt - 1; c++) {
			row.createCell();
		}
		for (int r = 1; r < rowCnt; r++) {
			row = tableOne.createRow();
		}
		setLocation(tableOne, 2);

		return tableOne;
	}

	public static XWPFParagraph createCursorParagraph(XWPFParagraph prgh) {

		XmlCursor cursor = prgh.getCTP().newCursor();
		XWPFParagraph prghNew = prgh.getBody().insertNewParagraph(cursor);
		return prghNew;
	}

	public static void setCellValueString(XWPFTable table, int row, int col,
			String value) {
		XWPFTableCell cell = table.getRow(row).getCell(col);
		XWPFParagraph pag;
		if(cell.getParagraphs().size()>0){
			pag = cell.getParagraphs().get(0);
		}else{
			pag = cell.addParagraph();
		}
		
		pag.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run ;
		if(pag.getRuns().size()>0){
			run = pag.getRuns().get(0);
		}else{
			run = pag.createRun();
		}		
		run.setText(value);
	}

	public static void setCellValueDouble(XWPFTable table, int row, int col,
			double value, int xiaoShuDian) {
		double v = new BigDecimal(value).setScale(xiaoShuDian,
				BigDecimal.ROUND_HALF_UP).doubleValue();
		setCellValueString(table, row, col, String.valueOf(v));
	}

	public static void setCellValueDouble(XWPFTable table, int row, int col,
			double value) {
		// double v = new BigDecimal(value).setScale(2,
		// BigDecimal.ROUND_HALF_UP).doubleValue();
		// setCellValueString(table,row,col,String.valueOf(v) );
		setCellValueDouble(table, row, col, value, 2);
	}

	public static void setTableColmnWidth(XWPFTable table, int[] widthArray) {

		// 设置指定宽度
		CTTbl ttbl = table.getCTTbl();
		CTTblGrid tblGrid = ttbl.addNewTblGrid();
		for (int i : widthArray) {
			CTTblGridCol gridCol = tblGrid.addNewGridCol();
			gridCol.setW(new BigInteger(i + ""));
		}
	}

	public static void setLocation(XWPFTable xwpfTable, int location) {
		CTTbl cttbl = xwpfTable.getCTTbl();
		CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl
				.getTblPr();
		CTJc cTJc = tblpr.addNewJc();
		cTJc.setVal(STJc.Enum.forInt(location));
	}
}
