package com.ss.poi.test;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.ss.poi.util.ReportFileUtil;
import com.ss.poi.util.OperateTable;
import com.ss.poi.util.WriteText;

/**
 * @Description
 */
public class CreateTableParaImp {
	
	/*public static void main(String[] args) throws Exception {
		XWPFDocument doc = new CreateTableParaImp().addFrame("这是第一段落介绍性文字！",
				new String[] { "1111111", "2", "3", "4", "5" });
		 XWPFDocument doc = CreateWordDocument.createNew();
		 OperateTable.createSpecifyTable(doc, 3, 3);
		 
		new CreateTableParaImp().insertData(doc, new String[] { "a","b","b", "c"  });
		CreateWordDocument.output(doc, "F:/a.docx");
	}*/
	
	/*public void void (string data){
		1\json to bean
		2\create docx
		3\switch(bean.functype)
		4\save docx
		5\save html
		6\
	}*/
	

	public XWPFDocument addFrame(String content, String[] tableTitle)
			throws Exception {
		XWPFDocument doc = ReportFileUtil.openDocx(null);
		WriteText.writeForPara(doc, content, 0);
		XWPFTable table = OperateTable.createAutoTable(doc);
		OperateTable.createTableTitle(table, tableTitle);
		return doc;
	}

	public void insertData(XWPFDocument doc, String[] content) throws Exception {
		XWPFTable table = doc.getTables().get(0);
		// 获取表格行数
		int tableRows = table.getRows().size();
		// 获取表格列数
		int tableCells = table.getRows().get(0).getTableCells().size();
		//内容占据单元格数量
		int contents = content.length;
		System.out.println("内容占据单元格数量:"+contents);
		System.out.println("获取表格行数：" + tableRows);
		System.out.println("获取表格列数：" + tableCells);
		if (tableRows > 1) {
			for (int i = 0; i < tableRows - 1; i++) {
				for (int j = 0; j < tableCells; j++) {
					XWPFTableCell tableCell = OperateTable.getTableCell(table,
							i + 1, j);
					if(i * tableCells + j<contents){
						tableCell.setText(content[i * tableCells + j]);
					}
				}
			}
		} else {
			/*根据内容获取行数*/
			//内容占据行数
			int countRows = (tableCells + contents - 1) / tableCells;
			for (int i = 0; i < countRows; i++) {
				XWPFTableRow row = table.createRow();
					for (int j = 0; j < contents; j++) {
						row.getCell(j).setText(content[i * tableCells + j]);
					}
			}
		}
	}

}
