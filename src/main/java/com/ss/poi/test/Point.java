package com.ss.poi.test;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.geom.Ellipse2D;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardXYItemLabelGenerator;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYItemRenderer;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.chart.title.LegendTitle;
import org.jfree.data.xy.DefaultXYDataset;
import org.jfree.data.xy.XYDataset;
import org.jfree.ui.TextAnchor;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.reflect.TypeToken;
import com.ss.poi.entity.TestResultData;
import com.ss.poi.util.GsonUtil;

/**
 * @Description
 */
public class Point {
	public static boolean createPointImg(String title,double xscale,double xposition,double[] yscale,double[] yoffset,List<List<Double>> listValue,List<String> lineNames,String filelocation) throws IOException {
		long atime = System.currentTimeMillis();
		// 步骤1：创建CategoryDataset对象（准备数据）
				XYDataset dataset = createxydataset(listValue,lineNames);
		long btime = System.currentTimeMillis();
		System.out.println("创建数据耗时" + (btime - atime));
		// 步骤2：根据Dataset 生成JFreeChart对象，以及做相应的设置
				JFreeChart freeChart = createPoint(dataset);
		long ctime = System.currentTimeMillis();
		System.out.println("绘制图表耗时" + (ctime - btime));
		// 步骤3：将JFreeChart对象输出到文件，Servlet输出流等
				ChartTest.saveAsFile(freeChart, filelocation, 1920, 1080);// 4
		long dtime = System.currentTimeMillis();
		System.out.println("输出文件耗时" + (dtime - ctime));
		System.out.println("总耗时" + (System.currentTimeMillis() - atime));
		return true;
	}

	public static void main(String[] args) throws IOException {
		String filepath = "e:/diyakai.data";
		String str = GsonUtil.readFile2String(filepath);
		TestResultData data = GsonUtil.jsonToObject(str, TestResultData.class);
		JsonArray array = data.getTestExportData().getAsJsonArray("waveData");// 1

		Gson gson = new Gson();
		List<Double> fromJson = gson.fromJson(array.toString(),
				new TypeToken<List<Double>>() {
				}.getType());
		ArrayList<List<Double>> list = new ArrayList<List<Double>>();
		list.add(fromJson);
		List<String> list2 = new ArrayList<String>();
		list2.add("a");
		long atime = System.currentTimeMillis();
		// 步骤1：创建CategoryDataset对象（准备数据）
		XYDataset dataset = createxydataset(list,list2);
		long btime = System.currentTimeMillis();
		System.out.println("创建数据耗时" + (btime - atime));
		// 步骤2：根据Dataset 生成JFreeChart对象，以及做相应的设置
		JFreeChart freeChart = createPoint(dataset);
		long ctime = System.currentTimeMillis();
		System.out.println("绘制图表耗时" + (ctime - btime));
		// 步骤3：将JFreeChart对象输出到文件，Servlet输出流等
		ChartTest.saveAsFile(freeChart, "E:\\line.jpg", 1920, 1080);// 4
		long dtime = System.currentTimeMillis();
		System.out.println("输出文件耗时" + (dtime - ctime));
		System.out.println("总耗时" + (System.currentTimeMillis() - atime));
	}

	public static JFreeChart createPoint(XYDataset xydataset) {

		JFreeChart jfreechart = ChartFactory.createScatterPlot("1", "2", "3",
				xydataset, PlotOrientation.VERTICAL, true, false, false);

		jfreechart.setBackgroundPaint(Color.white);
		jfreechart.setBorderPaint(Color.GREEN);
		jfreechart.setBorderStroke(new BasicStroke(1.5f));
		XYPlot xyplot = (XYPlot) jfreechart.getPlot();
		// xyplot.setNoDataMessage(nobloodData);
		// xyplot.setNoDataMessageFont(new Font("",Font.BOLD,14));
		// xyplot.setNoDataMessagePaint(new Color(87,149,117));
		 // 没有数据时显示的消息
		xyplot.setNoDataMessage("没有相关统计数据");
		xyplot.setNoDataMessageFont(new Font("黑体", Font.CENTER_BASELINE, 16));
		xyplot.setNoDataMessagePaint(Color.RED);
		
		XYLineAndShapeRenderer xylineandshaperenderer=new XYLineAndShapeRenderer();
//		   xylineandshaperenderer.setSeriesPaint(0, Color.green);//设置第一条曲线颜色
//		   xylineandshaperenderer.setSeriesPaint(1, Color.red);  //设置第二条曲线颜色   
//		   xylineandshaperenderer.setSeriesPaint(2, Color.blue);//设置第三条曲线颜色
//		   xylineandshaperenderer.setSeriesPaint(3, Color.orange);//设置第四条曲线颜色
		    
		   xylineandshaperenderer.setSeriesShape(0,new Ellipse2D.Double(1, 1, 0.0001, 0.0001));//设置第1条曲线数据点的图形
		   xylineandshaperenderer.setSeriesShape(1,new Ellipse2D.Double(1, 1, 0.0001, 0.0001));//设置第2条曲线数据点的图形
		   xylineandshaperenderer.setSeriesShape(2,new Ellipse2D.Double(1, 1, 0.0001, 0.0001));//设置第3条曲线数据点的图形
		   xylineandshaperenderer.setSeriesShape(3,new Ellipse2D.Double(1, 1, 0.0001, 0.0001));//设置第4条曲线数据点的图形
	/*	   xylineandshaperenderer.setSeriesOutlinePaint(0,Color.black);//设置第一条曲线数据点画图型的颜色
		    
		   xylineandshaperenderer.setSeriesFillPaint(0,Color.cyan);//设置第一条曲线数据点填充色
		   xylineandshaperenderer.setSeriesShapesVisible(0,true);//第一条线数据点可见
		   xylineandshaperenderer.setLinesVisible(false);//连线不可见
		    
		   xylineandshaperenderer.setUseOutlinePaint(false);//设置是否画曲线数据点的轮廓图形
		   xylineandshaperenderer.setUseFillPaint(false);    //设置是否填充曲线数据点 
		   xylineandshaperenderer.setDrawSeriesLineAsPath(true);
		   xylineandshaperenderer.setBaseLinesVisible(true);*/
		   
		   xyplot.setRenderer(xylineandshaperenderer);
		    
		   /*//设置曲线显示各数据点的值
		   XYItemRenderer xyitem = xyplot.getRenderer();
		   xyitem.setBaseItemLabelsVisible(true);
		   xyitem.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_LEFT));
		   //下面三句是对设置折线图数据标示的关键代码
		   xyitem.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
		   xyitem.setBaseItemLabelFont(new Font("Dialog", 1, 14));
		   xyplot.setRenderer(xyitem);*/
		   //隐藏图例
		   LegendTitle legend = jfreechart.getLegend();
		   legend.setVisible(false);

		//网格线设置
		xyplot.setBackgroundPaint(new Color(255, 253, 246));
		ValueAxis vaaxis = xyplot.getDomainAxis();
		vaaxis.setAxisLineStroke(new BasicStroke(1.5f));

		ValueAxis va = xyplot.getDomainAxis(0);
		va.setAxisLineStroke(new BasicStroke(1.5f));

		va.setAxisLineStroke(new BasicStroke(1.5f));// 坐标轴粗细
		va.setAxisLinePaint(new Color(215, 215, 215));// 坐标轴颜色
		xyplot.setOutlineStroke(new BasicStroke(1.5f));// 边框粗细
		va.setLabelPaint(new Color(10, 10, 10));// 坐标轴标题颜色
		va.setTickLabelPaint(new Color(102, 102, 102));// 坐标轴标尺值颜色
		ValueAxis axis = xyplot.getRangeAxis();
		axis.setAxisLineStroke(new BasicStroke(1.5f));
		//x轴是否可见
		//xyplot.getDomainAxis().setVisible(false);

		//NumberAxis rangeAxis = (NumberAxis) xyplot.getRangeAxis();
		// 数据轴精度
//		rangeAxis.setTickUnit(new NumberTickUnit(1));
//		rangeAxis.setLowerBound(-1);
//		rangeAxis.setUpperBound(1);

		return jfreechart;

	}
	

	public static XYDataset createxydataset(List<List<Double>> xydatalist,List<String> caption) {
		int dataCount = xydatalist.size();
		System.out.println("dataCount:"+dataCount);
		DefaultXYDataset xydataset = new DefaultXYDataset();
		for (int i = 0; i < dataCount; i++) {
			int dataSize = xydatalist.get(i).size();
			System.out.println("dataSize:"+dataSize);
			double[][] datas = new double[2][dataSize];
			for (int j = 0; j < dataSize; j++) {
				Double sys = xydatalist.get(i).get(j);
				datas[0][j] = j;
				datas[1][j] = sys;
			}
			xydataset.addSeries(caption.get(i), datas);
		}
		return xydataset;
	}
}
