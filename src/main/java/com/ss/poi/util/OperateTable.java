package com.ss.poi.util;

import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

/**
 * @Description
 */
public class OperateTable {
	/**
	 * 在指定word内生成指定行列表格，默认居中
	 * 
	 * @param doc
	 * @param rows
	 * @param cells
	 * @return XWPFTable
	 */
	public static XWPFTable createSpecifyTable(XWPFDocument doc, int rows, int cols) throws Exception {
		doc.createParagraph();
		XWPFTable table = doc.createTable(rows, cols);
		setLocation(table, 2);
		return table;
	}

	/**
	 * 在指定word内生成行列自动分割的表格，默认居中
	 * 
	 * @param doc
	 * @return XWPFTable
	 * @throws Exception
	 */
	public static XWPFTable createAutoTable(XWPFDocument doc) throws Exception {
		doc.createParagraph();
		XWPFTable table = doc.createTable();
		CTTblWidth infoTableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
		infoTableWidth.setType(STTblWidth.DXA);
		setLocation(table, 2);
		return table;
	}

	/**
	 * 在指定word内生成行列自动分割的表格，默认居中
	 * 
	 * @param doc
	 * @return XWPFTable
	 * @throws Exception
	 */
	public static void main(String[] args) {
		System.out.println(14 % 7);
	}

	/**
	 * 创建折线图附属点的表格，默认居中
	 * 
	 * @param doc
	 * @return XWPFTable
	 * @throws Exception
	 */
	public static void createLineChartTable(XWPFParagraph para, JsonObject array) throws Exception {
		XWPFTable table = null;
		if (array != null) {
			JsonArray chartArray = array.getAsJsonArray("CVs");
			if (chartArray != null & chartArray.size() > 0) {
				table = POITableUtil.createCursorTable(para, ((chartArray.size() + 2) / 3) + 1, 9);
				int tableHeanderIndex = 0;
				// 设置表头
				while (tableHeanderIndex < 3) {
					POITableUtil.setCellValueString(table, 0, 3 * tableHeanderIndex, "序号");
					POITableUtil.setCellValueString(table, 0, 3 * tableHeanderIndex + 1, "电阻");
					POITableUtil.setCellValueString(table, 0, 3 * tableHeanderIndex + 2, "电压");
					tableHeanderIndex++;
				}
				int xuHaoIndex = 0;
				int dianZuIndex = 0;
				int dianYaIndex = 0;
				for (int i = 0; i < chartArray.size(); i++) {
					String[] asd = chartArray.get(i).getAsString().split(",");
					if (i == 0 || i % 3 == 0) {
						xuHaoIndex = 0;
						dianZuIndex = 1;
						dianYaIndex = 2;
					} else if (i % 3 == 1) {
						xuHaoIndex = 3;
						dianZuIndex = 4;
						dianYaIndex = 5;
					} else if (i % 3 == 2) {
						xuHaoIndex = 6;
						dianZuIndex = 7;
						dianYaIndex = 8;
					}
					POITableUtil.setCellValueString(table, (i + 3) / 3, xuHaoIndex, (i + 1) + "");
					POITableUtil.setCellValueDouble(table, (i + 3) / 3, dianZuIndex, Double.parseDouble(asd[0]));
					POITableUtil.setCellValueDouble(table, (i + 3) / 3, dianYaIndex, Double.parseDouble(asd[1]));
				}
			}
		}
		POITableUtil.setTableColmnWidth(table, new int[] { 800, 1200, 1200, 800, 1200, 1200, 800, 1200, 1200 });
	}
	/**
	 * 创建折线图附属点的表格，默认居中
	 * 
	 * @param doc
	 * @return XWPFTable
	 * @throws Exception
	 */
	public static void createLineChartTable2(XWPFParagraph para, JsonObject array) throws Exception {
		XmlCursor cursor = para.getCTP().newCursor();
		XWPFTable table = para.getBody().insertNewTbl(cursor);
		if (array != null) {
			JsonArray chartArray = array.getAsJsonArray("CVs");
			XWPFTableRow row = null;
			if (chartArray != null & chartArray.size() > 0) {
				for (int i = 1; i < chartArray.size() + 1; i++) {
					if (i < 8) {
						if (i == 1) {
							row = table.getRow(0);
							row.getCell(i - 1).setText(chartArray.get(i - 1).getAsString());
						} else {
							row.createCell().setText(chartArray.get(i - 1).getAsString());
						}
					} else {
						if (i % 7 == 1) {
							row = table.createRow();
						}
						row.getCell(i % 7 - 1 == -1 ? 6 : i % 7 - 1).setText(chartArray.get(i - 1).getAsString());
					}
				}
			}
		}
		setLocation(table, 2);
	}

	/**
	 * 在指定表格生成新表头
	 * 
	 * @param table
	 * @param content
	 * @return
	 */
	public static boolean createTableTitle(XWPFTable table, String[] content) {
		try {
			for (int i = 0; i < content.length; i++) {
				if (i == 0) {
					table.getRows().get(0).getCell(0).setText(content[i]);
				} else {
					try {
						table.getRows().get(0).getCell(i).setText(content[i]);
					} catch (Exception e) {
						table.getRows().get(0).addNewTableCell().setText(content[i]);
					}
				}
			}
		} catch (Exception e) {
			System.out.println("OperateTable_setLocation出错！");
			return false;
		}
		return true;
	}

	/**
	 * 获取单元格对象
	 * 
	 * @return
	 * @throws Exception
	 */
	public static XWPFTableCell getTableCell(XWPFTable table, int rows, int cols) throws Exception {
		XWPFTableRow row = table.getRow(rows);
		return row.getCell(cols);
	}

	/**
	 * 获取行对象
	 * 
	 * @return
	 * @throws Exception
	 */
	public static XWPFTable getTableRow() throws Exception {

		return null;
	}

	/**
	 * 设置表格位置
	 * 
	 * @param xwpfTable
	 * @param location  表格在段落的位置：1居左，2居中，3居右，4两端对齐
	 */
	public static boolean setLocation(XWPFTable xwpfTable, int location) {
		try {
			CTTbl cttbl = xwpfTable.getCTTbl();
			CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl.getTblPr();
			CTJc cTJc = tblpr.addNewJc();
			cTJc.setVal(STJc.Enum.forInt(location));
		} catch (Exception e) {
			System.out.println("OperateTable_setLocation出错！");
			return false;
		}
		return true;
	}

	/**
	 * 设置表格内所有行列宽高
	 * 
	 * @param table
	 * @param widthArray  行宽数组
	 * @param HeightArray 列高数组
	 * @throws Exception
	 */
	public static boolean setStyle(XWPFTable table, int[] widthArray, int[] HeightArray) {
		try {
			// 设置指定宽度
			CTTbl ttbl = table.getCTTbl();
			CTTblGrid tblGrid = ttbl.addNewTblGrid();
			for (int i : widthArray) {
				CTTblGridCol gridCol = tblGrid.addNewGridCol();
				gridCol.setW(new BigInteger(i + ""));
			}
			// 设置指定高度
			for (int i = 0; i < HeightArray.length; i++) {
				table.getRow(i).setHeight(HeightArray[i]);
			}
		} catch (Exception e) {
			System.out.println("OperateTable_setLocation出错！");
			return false;
		}
		return true;
	}

}
