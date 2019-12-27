package com.ss.poi.util;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryAnchor;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.data.xy.DefaultXYDataset;
import org.jfree.data.xy.XYDataset;

import java.awt.*;
import java.awt.geom.Ellipse2D;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class LineChartUtil {
    /**
     * 创建JFreeChart Line Chart（折线图）
     *
     * @throws IOException
     */
    public static String createChart(JsonObject testExportData,String xLegend,String yLegend,String lineName){
        // 步骤1：创建CategoryDataset对象（准备数据）
        XYDataset dataset = createxydataset(testExportData,lineName);
        // 步骤2：根据Dataset 生成JFreeChart对象，以及做相应的设置
        JFreeChart freeChart = createJFreeChart(dataset,xLegend,yLegend);
        // 步骤3：将JFreeChart对象输出到文件，Servlet输出流等
        String outFilePath = saveAsFile(freeChart, 1100, 700);
        return outFilePath;
    }

    // 保存为文件
    public static String saveAsFile(JFreeChart chart,int weight, int height) {
        FileOutputStream out = null;
        String imgPath = ReportFileUtil.getTempImageLocation();
        try {
            File outFile = new File(imgPath);
            if (!outFile.getParentFile().exists()) {
                outFile.getParentFile().mkdirs();
            }
            out = new FileOutputStream(imgPath);
            // 保存为PNG
            // ChartUtilities.writeChartAsPNG(out, chart, 600, 400);
            // 保存为JPEG 图片质量，0~1
            ChartUtilities.writeChartAsJPEG(out, 1, chart, weight, height);
//            FileOutputStream out2 = new FileOutputStream("e:/"+ System.currentTimeMillis() + ".jpg");
//            ChartUtilities.writeChartAsJPEG(out2, 1, chart, weight, height);
//            out2.flush();
//            out2.close();
            out.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    // do nothing
                }
            }
            return imgPath;
        }
    }

    // 根据CategoryDataset创建JFreeChart对象
    private static JFreeChart createJFreeChart(XYDataset dataset,String xLegend,String yLegend) {
        // 创建JFreeChart对象：ChartFactory.createLineChart
        JFreeChart jfreechart = ChartFactory.createScatterPlot("", // 标题
        		xLegend, // categoryAxisLabel （category轴，横轴，X轴标签）
        		yLegend, // valueAxisLabel（value轴，纵轴，Y轴的标签）
                dataset, // dataset
                PlotOrientation.VERTICAL, true, // legend
                false, // tooltips
                false); // URLs
		//背景色
        //jfreechart.setBackgroundPaint(Color.white);
        //jfreechart.setBorderPaint(Color.GREEN);
        jfreechart.setBorderStroke(new BasicStroke(1.5f));
        XYPlot xyplot = (XYPlot) jfreechart.getPlot();
        
        //网格线
        //x
        //xyplot.setDomainGridlineStroke(new BasicStroke());
        xyplot.setDomainGridlinesVisible(true);
        xyplot.setDomainGridlinePaint(new Color(130,130,130));
        //y
        //xyplot.setRangeGridlineStroke(new BasicStroke());
        xyplot.setRangeGridlinesVisible(true);
        xyplot.setRangeGridlinePaint(new Color(130,130,130));

        // xyplot.setNoDataMessage(nobloodData);
        // xyplot.setNoDataMessageFont(new Font("",Font.BOLD,14));
        // xyplot.setNoDataMessagePaint(new Color(87,149,117));
        // 没有数据时显示的消息
        xyplot.setNoDataMessage("无相关数据");
        xyplot.setNoDataMessageFont(new Font("黑体", Font.CENTER_BASELINE, 16));
        xyplot.setNoDataMessagePaint(Color.RED);

        XYLineAndShapeRenderer xylineandshaperenderer = new XYLineAndShapeRenderer();
        xylineandshaperenderer.setSeriesPaint(0, Color.blue);// 设置第一条曲线颜色
        xylineandshaperenderer.setSeriesPaint(1, Color.red); // 设置第二条曲线颜色
        xylineandshaperenderer.setSeriesPaint(2, Color.green);// 设置第三条曲线颜色
        xylineandshaperenderer.setSeriesPaint(3, Color.orange);// 设置第四条曲线颜色

        /* xylineandshaperenderer.setSeriesShape(0, new Ellipse2D.Double(1, 1,
                0.0001, 0.0001));// 设置第1条曲线数据点的图形
        xylineandshaperenderer.setSeriesShape(1, new Ellipse2D.Double(1, 1,
                0.0001, 0.0001));// 设置第2条曲线数据点的图形
        xylineandshaperenderer.setSeriesShape(2, new Ellipse2D.Double(1, 1,
                0.0001, 0.0001));// 设置第3条曲线数据点的图形
        xylineandshaperenderer.setSeriesShape(3, new Ellipse2D.Double(1, 1,
                0.0001, 0.0001));// 设置第4条曲线数据点的图形
        
         * xylineandshaperenderer.setSeriesOutlinePaint(0,Color.black);//
     	   * 设置第一条曲线数据点画图型的颜色
         *
         * xylineandshaperenderer.setSeriesFillPaint(0,Color.cyan);//设置第一条曲线数据点填充色
         * xylineandshaperenderer.setSeriesShapesVisible(0,true);//第一条线数据点可见
         * xylineandshaperenderer.setLinesVisible(false);//连线不可见
         *
         * xylineandshaperenderer.setUseOutlinePaint(false);//设置是否画曲线数据点的轮廓图形
         * xylineandshaperenderer.setUseFillPaint(false); //设置是否填充曲线数据点
         * xylineandshaperenderer.setDrawSeriesLineAsPath(true);
         * xylineandshaperenderer.setBaseLinesVisible(true);
         */
        xylineandshaperenderer.setSeriesShapesVisible(0,false);//第一条线数据点不可见
        xyplot.setRenderer(xylineandshaperenderer);

        // 设置曲线显示各数据点的值
        /*XYItemRenderer xyitem = xyplot.getRenderer();
        xyitem.setBaseItemLabelsVisible(true);
        xyitem.setBasePositiveItemLabelPosition(new ItemLabelPosition(
                ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_CENTER));
        // 下面三句是对设置折线图数据标示的关键代码
        xyitem.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
        xyitem.setBaseItemLabelFont(new Font("Dialog", 1, 10));
        xyplot.setRenderer(xyitem);*/
        // 隐藏图例
		/*LegendTitle legend = jfreechart.getLegend();
		legend.setVisible(false);*/

        xyplot.setBackgroundPaint(new Color(255, 253, 255));
        ValueAxis vaaxis = xyplot.getDomainAxis();
        vaaxis.setAxisLineStroke(new BasicStroke(2f));

        NumberAxis numberaxis = (NumberAxis) xyplot.getRangeAxis();
        numberaxis.setAxisLineStroke(new BasicStroke(2f));
        /*------设置X轴坐标上的文字-----------*/
        vaaxis.setTickLabelFont(new Font("sans-serif", Font.PLAIN, 12));
        /*------设置X轴的标题文字------------*/
        vaaxis.setLabelFont(new Font("黑体", Font.PLAIN, 15));
        /*------设置Y轴坐标上的文字-----------*/
        numberaxis.setTickLabelFont(new Font("sans-serif", Font.PLAIN, 12));
        /*------设置Y轴的标题文字------------*/
        numberaxis.setLabelFont(new Font("黑体", Font.PLAIN, 15));
        /*------这句代码解决了底部汉字乱码的问题-----------*/
        jfreechart.getLegend().setItemFont(new Font("黑体", Font.PLAIN, 15));

        // xy轴颜色及字体
        //x
        ValueAxis va = xyplot.getDomainAxis();
        va.setAxisLineStroke(new BasicStroke(1.5f));// 坐标轴粗细
        va.setAxisLinePaint(new Color(0,0,0));// 坐标轴颜色
        xyplot.setOutlineStroke(new BasicStroke(1.5f));// 边框粗细
        va.setLabelPaint(new Color(0,0,0));// 坐标轴标题颜色
        va.setTickLabelPaint(new Color(0,0,0));// 坐标轴标尺值颜色
        //y
        ValueAxis axis = xyplot.getRangeAxis();
        axis.setAxisLineStroke(new BasicStroke(1.5f));// 坐标轴粗细
        axis.setAxisLinePaint(new Color(0,0,0));// 坐标轴颜色
        axis.setLabelPaint(new Color(0,0,0));// 坐标轴标题颜色
        axis.setTickLabelPaint(new Color(0,0,0));// 坐标轴标尺值颜色
        // x轴是否可见
        xyplot.getDomainAxis().setVisible(true);

        // 数据轴精度
        // NumberAxis rangeAxis = (NumberAxis) xyplot.getRangeAxis();
        // rangeAxis.setTickUnit(new NumberTickUnit(1));
        // rangeAxis.setLowerBound(-1);
        // rangeAxis.setUpperBound(1);
        return jfreechart;
    }

    public static XYDataset createxydataset(JsonObject testExportData,String lineName){
        DefaultXYDataset xydataset = new DefaultXYDataset();
        if(testExportData!=null){
            JsonArray array = testExportData.getAsJsonArray("CVs");
            System.out.println("折线图点量："+ array.size());
            double[][] datas = new double[2][array.size()];
            for (int i = 0; i < array.size(); i++) {
                //double[][] datasa= {{1,2,3},{4.6,4.8,6.8}};
                String[] asd = array.get(i).getAsString().split(",");
                datas[0][i] = Double.parseDouble(asd[0]);
                datas[1][i] = Double.parseDouble(asd[1]);
            }
            xydataset.addSeries(lineName, datas);
        }
        return xydataset;
    }
}
