package com.ss.poi.util;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.List;

//m*9的表格
public class TableHeaderController  {
    public static void main(String[] args)throws Exception {
        XWPFDocument document = new XWPFDocument();
        System.out.println("文档对象创建成功:"+document);
        //保存文档到磁盘中
        FileOutputStream out = new FileOutputStream(new File("F:\\表头.doc"));
        //创建一个m*9的表格
        XWPFTable table = document.createTable(6, 9);
        //添加数据，内容居中
        XWPFTableRow row = table.getRow(0);
        //设置表头
        for (int colIndex = 0; colIndex < 9; colIndex=colIndex+3) {
            XWPFTableCell cell0 = row.getCell(colIndex);
            cell0.addParagraph().createRun().setText("序号");
            XWPFTableCell cell1 = row.getCell(colIndex+1);
            cell1.addParagraph().createRun().setText("电阻");
            XWPFTableCell cell2 = row.getCell(colIndex+2);
            cell2.addParagraph().createRun().setText("电压");

        }
        //调用方法，进行样式设置
        TableHeaderController tableHeaderController = new TableHeaderController();
        //设置表格的宽度和内容居中位置
        tableHeaderController.setTableWidthAndHAlign(table, "9072", STJc.CENTER);
        //写入数据
        int index=0;
        for (int rowIndex = 1; rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow rowTemp = table.getRow(rowIndex);
            for (int colIndex = 0; colIndex < rowTemp.getTableCells().size(); colIndex++) {
                XWPFTableCell cell = rowTemp.getCell(colIndex);
                if(colIndex%3==0){
                    index++;
                    cell.addParagraph().createRun().setText(""+index);
                }else{
                    cell.addParagraph().createRun().setText("cell"+rowIndex+colIndex);
                }
            }
        }
        document.write(out);
        out.close();
        document.close();
        System.out.println("表头样式创建成功");
    }

    //获取表格中的某一个格子
    public XWPFTableCell getCell( XWPFTable table,int row, int cell){

        XWPFTableRow tableRow = table.getRow(row);
        XWPFTableCell cell1 = tableRow.getCell(cell);
        return cell1;

    }

    //设置表格的行高
    public void setTableHeight(XWPFTable infoTable, int heigth, STVerticalJc.Enum vertical) {
        List<XWPFTableRow> rows = infoTable.getRows();
        for(XWPFTableRow row : rows) {
            CTTrPr trPr = row.getCtRow().addNewTrPr();
            CTHeight ht = trPr.addNewTrHeight();
            ht.setVal(BigInteger.valueOf(heigth));
            List<XWPFTableCell> cells = row.getTableCells();
            for(XWPFTableCell tableCell : cells ) {
                CTTcPr cttcpr = tableCell.getCTTc().addNewTcPr();
                cttcpr.addNewVAlign().setVal(vertical);
            }
        }
    }

    //设置表格的宽度和位置
    public void setTableWidthAndHAlign(XWPFTable table, String width, STJc.Enum enumValue) {
        CTTblPr tblPr = getTableCTTblPr(table);
        // 表格宽度
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        if (enumValue != null) {
            CTJc cTJc = tblPr.addNewJc();
            cTJc.setVal(enumValue);
        }
        // 设置宽度
        tblWidth.setW(new BigInteger(width));
        tblWidth.setType(STTblWidth.DXA);
    }

    //获取表格
    public CTTblPr getTableCTTblPr(XWPFTable table) {
        CTTbl ttbl = table.getCTTbl();
        // 表格属性
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        return tblPr;
    }
}


