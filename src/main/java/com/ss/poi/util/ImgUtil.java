package com.ss.poi.util;

import com.ss.poi.entity.LinePoint;
import com.ss.poi.test.ChartTest;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.axis.NumberTickUnit;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardXYItemLabelGenerator;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYItemRenderer;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.data.xy.DefaultXYDataset;
import org.jfree.data.xy.XYDataset;
import org.jfree.ui.TextAnchor;

import java.awt.*;
import java.awt.geom.Ellipse2D;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @Description
 */
public class ImgUtil {
	/**
	 * 在段落内插入图片
	 * 
	 * @param doc
	 * @param path
	 * @param location
	 *            图片在段落的位置：1居左，2居中，3居右，4两端对齐
	 * @throws Exception
	 */
	public static void insertForPara(XWPFParagraph para, String base64,
			int width, int height) throws Exception {
		XWPFRun run;
		para.setAlignment(ParagraphAlignment.CENTER);
		run = para.createRun();
		String imgPath = ReportFileUtil.byte64ToTempImg(base64);
		InputStream input = new FileInputStream(imgPath);
		run.addPicture(input, XWPFDocument.PICTURE_TYPE_JPEG, imgPath,
				Units.toEMU(width), Units.toEMU(height));
		input.close();
		ReportFileUtil.delTempFile(imgPath);
	}

	/**
	 * 散点图
	 * 
	 * @param title
	 * @param xscale
	 * @param xposition
	 * @param yscale
	 * @param yoffset
	 * @param listValue
	 *            数据集合
	 * @param lineNames
	 * @param filelocation
	 *            保存文件路径，包含文件名及后缀
	 * @return boolean
	 * @throws IOException
	 */
	public static String createPointImg(String title, double xscale,
			double xposition, double[] yscale, double[] yoffset,
			List<List<Double>> listValue, List<String> lineNames)
			throws IOException {
		String imgPath = ReportFileUtil.getTempImageLocation();
		new ImgUtil().createPointImg(title, xscale, xposition, yscale, yoffset,
				listValue, lineNames, imgPath);
		return imgPath;
	}

	public boolean createPointImg(String title, double xscale,
			double xposition, double[] yscale, double[] yoffset,
			List<List<Double>> listValue, List<String> lineNames,
			String filelocation) throws IOException {
		long atime = System.currentTimeMillis();
		// 步骤1：创建CategoryDataset对象（准备数据）
		XYDataset dataset = createxydatasetNoXAxis(listValue, lineNames);
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

	/**
	 * 绘制无x轴的连线图
	 * 
	 * @param listValue
	 *            数据集合
	 * @param filelocation
	 *            保存文件路径，包含文件名及后缀
	 * @return boolean
	 * @throws IOException
	 */
	public static String createLineImgNoXAxis(List<Double> ylistValue ,
			List<String> lineNameList, String xAxisName, String yAxisName)
			throws IOException {
		String imgPath = ReportFileUtil.getTempImageLocation();
		new ImgUtil().createLineImgNoXAxis(ylistValue, imgPath,false,true,lineNameList,xAxisName,yAxisName);
		return imgPath;
	}
	public static String createLineImgHasXAxisNoLabel(List<LinePoint> pointLIst,String xAxisName,
			String yAxisName,String lineName)
			throws IOException {
		String imgPath = ReportFileUtil.getTempImageLocation();
		ArrayList<String> lineNameList = new ArrayList<String>();
		ArrayList<Double> xList = new ArrayList<Double>();
		ArrayList<Double> yList = new ArrayList<Double>();
		lineNameList.add(lineName);

		if(pointLIst.size()>0){
			for (int i = 0; i < pointLIst.size(); i++) {
				xList.add(pointLIst.get(i).getX());
				yList.add(pointLIst.get(i).getY());
			}
		}
		new ImgUtil().createLineImgHasXAxis(xList,yList, imgPath,true,false,xAxisName,yAxisName,lineNameList);
		return imgPath;
	}
	public static String createLineImgHasXAxisNoLabel(List<Double> xlistValue,List<Double> ylistValue,
			String xAxisName,String yAxisName,List<String> lineNameList)
			throws IOException {
		String imgPath = ReportFileUtil.getTempImageLocation();
		new ImgUtil().createLineImgHasXAxis(xlistValue,ylistValue, imgPath,true,false,xAxisName,yAxisName,lineNameList);
		return imgPath;
	}

	public boolean createLineImgHasXAxis(List<Double> xlistValue,List<Double> ylistValue, String filelocation,boolean xAxisVisible
			,boolean labelVisible,String xAxisName,String yAxisName,List<String> lineNameList) throws IOException{
		ArrayList<List<Double>> xlist = new ArrayList<List<Double>>();
		xlist.add(xlistValue);
		ArrayList<List<Double>> ylist = new ArrayList<List<Double>>();
		ylist.add(ylistValue);
	
		XYDataset dataset = createxydatasetHasXAxis(xlist, ylist,lineNameList);
		createLineImg(dataset, filelocation,xAxisVisible,labelVisible,xAxisName,yAxisName);
		return true;
	}
	public boolean createLineImgNoXAxis(List<Double> yValue, String filelocation,boolean xAxisVisible,
			boolean labelVisible,List<String> lineNameList,String xAxisName,String yAxisName)
			throws IOException {
		ArrayList<List<Double>> list = new ArrayList<List<Double>>();
		list.add(yValue);
		ArrayList<String> list2 = new ArrayList<String>();
		if(lineNameList.size()>0){
			for(String lineName:lineNameList ){
				list2.add(lineName);
			}
		}
		XYDataset dataset = createxydatasetNoXAxis(list, list2);
		createLineImg(dataset, filelocation,xAxisVisible,labelVisible,xAxisName,yAxisName);
		return true;
	}
	public boolean createLineImg(XYDataset dataset , String filelocation,boolean xAxisVisible,boolean labelVisible,
			String xAxisName,String yAxisName)
			throws IOException {
		long atime = System.currentTimeMillis();
		JFreeChart freeChart = null;
		// 步骤1：创建CategoryDataset对象（准备数据）
		/*ArrayList<List<Double>> list = new ArrayList<List<Double>>();
		list.add(listValue);
		ArrayList<String> list2 = new ArrayList<String>();
		list2.add("");
		XYDataset dataset = createxydatasetNoXAxis(list, list2);*/
		long btime = System.currentTimeMillis();
		System.out.println("创建数据耗时" + (btime - atime));
		// 步骤2：根据Dataset 生成JFreeChart对象，以及做相应的设置
		if(xAxisVisible){
			freeChart = createLineHasXAxis(dataset,labelVisible,xAxisName,yAxisName);
		}else{
			freeChart = createLineNoXAxis(dataset,labelVisible,xAxisName,yAxisName);
		}
		long ctime = System.currentTimeMillis();
		System.out.println("绘制图表耗时" + (ctime - btime));
		// 步骤3：将JFreeChart对象输出到文件，Servlet输出流等
		ChartTest.saveAsFile(freeChart, filelocation, 1920, 1080);// 4
		long dtime = System.currentTimeMillis();
		System.out.println("输出文件耗时" + (dtime - ctime));
		System.out.println("总耗时" + (System.currentTimeMillis() - atime));
		return true;
	}
	public JFreeChart createLineHasXAxis(XYDataset xydataset,boolean labelVisible,String xAxisName,String yAxisName) {
		return createLine(xydataset, true,labelVisible, xAxisName, yAxisName);
	}
	public JFreeChart createLineNoXAxis(XYDataset xydataset,boolean labelVisible,String xAxisName,String yAxisName ) {
		return createLine(xydataset, false,labelVisible, xAxisName, yAxisName);
	}
	public JFreeChart createLine(XYDataset xydataset,boolean xAxisVisible,boolean labelVisible,
			String xAxisName,String yAxisName) {

		JFreeChart jfreechart = ChartFactory.createScatterPlot("", xAxisName, yAxisName,
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

		XYLineAndShapeRenderer xylineandshaperenderer = new XYLineAndShapeRenderer();
		xylineandshaperenderer.setSeriesPaint(0, Color.green);// 设置第一条曲线颜色
		xylineandshaperenderer.setSeriesPaint(1, Color.red); // 设置第二条曲线颜色
		xylineandshaperenderer.setSeriesPaint(2, Color.blue);// 设置第三条曲线颜色
		xylineandshaperenderer.setSeriesPaint(3, Color.orange);// 设置第四条曲线颜色

		xylineandshaperenderer.setSeriesShape(0, new Ellipse2D.Double(1, 1,
				0.0001, 0.0001));// 设置第1条曲线数据点的图形
		xylineandshaperenderer.setSeriesShape(1, new Ellipse2D.Double(1, 1,
				0.0001, 0.0001));// 设置第2条曲线数据点的图形
		xylineandshaperenderer.setSeriesShape(2, new Ellipse2D.Double(1, 1,
				0.0001, 0.0001));// 设置第3条曲线数据点的图形
		xylineandshaperenderer.setSeriesShape(3, new Ellipse2D.Double(1, 1,
				0.0001, 0.0001));// 设置第4条曲线数据点的图形
		/*
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

		xyplot.setRenderer(xylineandshaperenderer);

		// 设置曲线显示各数据点的值
		XYItemRenderer xyitem = xyplot.getRenderer();
		xyitem.setBaseItemLabelsVisible(labelVisible);
		xyitem.setBasePositiveItemLabelPosition(new ItemLabelPosition(
				ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_CENTER));
		// 下面三句是对设置折线图数据标示的关键代码
		xyitem.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator());
		xyitem.setBaseItemLabelFont(new Font("Dialog", 1, 14));
		xyplot.setRenderer(xyitem);
		// 隐藏图例
		/*LegendTitle legend = jfreechart.getLegend();
		legend.setVisible(false);*/

		xyplot.setBackgroundPaint(new Color(255, 253, 246));
		ValueAxis vaaxis = xyplot.getDomainAxis();
		vaaxis.setAxisLineStroke(new BasicStroke(1.5f));
		NumberAxis numberaxis = (NumberAxis) xyplot.getRangeAxis();
		 
	    /*------设置X轴坐标上的文字-----------*/
		vaaxis.setTickLabelFont(new Font("sans-serif", Font.PLAIN, 12));
	    /*------设置X轴的标题文字------------*/
		vaaxis.setLabelFont(new Font("黑体", Font.PLAIN, 15));
	    /*------设置Y轴坐标上的文字-----------*/
		numberaxis.setTickLabelFont(new Font("黑体", Font.PLAIN, 12));
	    /*------设置Y轴的标题文字------------*/
		numberaxis.setLabelFont(new Font("黑体", Font.PLAIN, 15));
	    /*------这句代码解决了底部汉字乱码的问题-----------*/
	    jfreechart.getLegend().setItemFont(new Font("黑体", Font.PLAIN, 15));

		// 背景网格线
		/*
		 * ValueAxis va = xyplot.getDomainAxis(0); va.setAxisLineStroke(new
		 * BasicStroke(1.5f));
		 * 
		 * va.setAxisLineStroke(new BasicStroke(1.5f));// 坐标轴粗细
		 * va.setAxisLinePaint(new Color(215, 215, 215));// 坐标轴颜色
		 * xyplot.setOutlineStroke(new BasicStroke(1.5f));// 边框粗细
		 * va.setLabelPaint(new Color(10, 10, 10));// 坐标轴标题颜色
		 * va.setTickLabelPaint(new Color(102, 102, 102));// 坐标轴标尺值颜色 ValueAxis
		 * axis = xyplot.getRangeAxis(); axis.setAxisLineStroke(new
		 * BasicStroke(1.5f));
		 */
		// x轴是否可见
		xyplot.getDomainAxis().setVisible(xAxisVisible);

		// 数据轴精度
		// NumberAxis rangeAxis = (NumberAxis) xyplot.getRangeAxis();
		// rangeAxis.setTickUnit(new NumberTickUnit(1));
		// rangeAxis.setLowerBound(-1);
		// rangeAxis.setUpperBound(1);

		return jfreechart;

	}

	public JFreeChart createPoint(XYDataset xydataset) {

		JFreeChart jfreechart = ChartFactory.createScatterPlot("", "", "",
				xydataset, PlotOrientation.VERTICAL, true, false, false);

		jfreechart.setBackgroundPaint(Color.white);
		jfreechart.setBorderPaint(Color.GREEN);
		jfreechart.setBorderStroke(new BasicStroke(1.5f));
		XYPlot xyplot = (XYPlot) jfreechart.getPlot();
		xyplot.setNoDataMessage("No Data!");
		xyplot.setNoDataMessageFont(new Font("", Font.BOLD, 14));
		xyplot.setNoDataMessagePaint(new Color(87, 149, 117));
		XYItemRenderer lasp = xyplot.getRenderer();
		lasp.setSeriesStroke(0, new BasicStroke(5F));

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

		xyplot.getDomainAxis().setVisible(false);

		NumberAxis rangeAxis = (NumberAxis) xyplot.getRangeAxis();
		// 数据轴精度
		rangeAxis.setTickUnit(new NumberTickUnit(1));
		rangeAxis.setLowerBound(-1);
		rangeAxis.setUpperBound(1);

		return jfreechart;

	}

	public XYDataset createxydatasetHasXAxis(List<List<Double>> xdatalist,List<List<Double>> ydatalist,List<String> caption) {
		return createxydataset(ydatalist, xdatalist,caption);
		
	}
	public XYDataset createxydatasetNoXAxis(List<List<Double>> ydatalist,List<String> caption) {
		return createxydataset(ydatalist, new ArrayList<List<Double>>(),caption);
		
	}
	public XYDataset createxydataset(List<List<Double>> xdatalist,List<List<Double>> ydatalist,
			List<String> caption) {
		Double x;
		int dataCount = ydatalist.size();
		System.out.println("ydatalist:" + dataCount);
		System.out.println("xdatalist:" + xdatalist.size());
		DefaultXYDataset xydataset = new DefaultXYDataset();
		for (int i = 0; i < dataCount; i++) {
			int ydataSize = ydatalist.get(i).size();
			int xdataSize = xdatalist.get(i).size();
			System.out.println("dataSize:" + ydataSize);
			System.out.println("dataSize:" + xdatalist.get(i).size());
			double[][] datas = new double[2][ydataSize];
			for (int j = 0; j < ydataSize; j++) {
				if(xdataSize>0){
					x = xdatalist.get(i).get(j);
				}else{
					x=j+0.0;
				}
				Double y = ydatalist.get(i).get(j);
				datas[0][j] = y;
				datas[1][j] = x;
			}
			xydataset.addSeries(caption.get(i), datas);
		}
		return xydataset;
	}

	// 保存为文件
	public void saveAsFile(JFreeChart chart, String outputPath,
			int weight, int height) {
		FileOutputStream out = null;
		try {
			File outFile = new File(outputPath);
			if (!outFile.getParentFile().exists()) {
				outFile.getParentFile().mkdirs();
			}
			out = new FileOutputStream(outputPath);
			// 保存为PNG
			// ChartUtilities.writeChartAsPNG(out, chart, 600, 400);
			// 保存为JPEG 图片质量，0~1
			ChartUtilities.writeChartAsJPEG(out, 1, chart, weight, height);
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
		}
	}

}
