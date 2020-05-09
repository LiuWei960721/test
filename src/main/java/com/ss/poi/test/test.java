package com.ss.poi.test;

import com.ss.poi.controller.WordReport;
import com.ss.poi.util.GsonUtil;

/**
 * @创建人: Liu
 * @创建时间: 2020-04-29 16:58
 * @描述:
 */

public class test {
    public static void main(String[] args) throws Exception {
        String filepath = "D:/L/LiuWeiSS7/20200312-143650.data";
        String dataStr = GsonUtil.readFile2String(filepath);
     System.out.println(dataStr);
        WordReport.creatWordReport(dataStr);
    }

    /*public static void main(String[] args) {
        String number = "123.456";

        System.out.println(number.split(".")[1]);
    }*/

}
