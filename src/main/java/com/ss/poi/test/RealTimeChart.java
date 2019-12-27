package com.ss.poi.test;
import java.awt.BorderLayout;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JFrame;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.time.Millisecond;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;

import com.google.gson.JsonArray;
import com.ss.poi.entity.TestResultData;
import com.ss.poi.util.GsonUtil;


public class RealTimeChart extends ChartPanel  {
    private static TimeSeries timeSeries;
    private long value = 0;

    public RealTimeChart(String chartContent, String title, String yaxisName) {
        super(createChart(chartContent, title, yaxisName));
    }

    @SuppressWarnings("deprecation")
	public static JFreeChart createChart(String chartContent, String title, String yaxisName) {
        //创建时序图对象
        timeSeries = new TimeSeries(chartContent, Millisecond.class);
        TimeSeriesCollection timeseriescollection = new TimeSeriesCollection(timeSeries);
        JFreeChart jfreechart = ChartFactory.createTimeSeriesChart(title, "时间(秒)", yaxisName, timeseriescollection, true, true, false);
        XYPlot xyplot = jfreechart.getXYPlot();
        //纵坐标设定
        ValueAxis valueaxis = xyplot.getDomainAxis();
        //自动设置数据轴数据范围
        valueaxis.setAutoRange(true);
        //数据轴固定数据范围 30s
        valueaxis.setFixedAutoRange(30000D);

        valueaxis = xyplot.getRangeAxis();
        //valueaxis.setRange(0.0D,200D);

        return jfreechart;
    }

    // 保存为文件
    public static void saveAsFile(JFreeChart chart, String outputPath,
                                  int weight, int height) throws Exception {
        FileOutputStream out = null;
        try {
            File outFile = new File(outputPath);
            if (!outFile.getParentFile().exists()) {
                outFile.getParentFile().mkdirs();
            }
            out = new FileOutputStream(outputPath);
            // 保存为PNG
            ChartUtilities.writeChartAsPNG(out, chart, 1000, 1000);
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

		public static JsonArray randomNum2() throws Exception {
        String filepath = "e:/diyakai.data";
        String str = GsonUtil.readFile2String(filepath);
        TestResultData data = GsonUtil.jsonToObject(str, TestResultData.class);
        JsonArray array = data.getTestExportData().getAsJsonArray("waveData");
       /* String []ss = str.split(",");
		double []random = new double[ss.length];
        for(int i=0;i<=ss.length;i++){
            random[i]=Double.parseDouble(ss[i]);
            //System.out.println(random[i]);
        }*/
        return array;
    }

    public static void run1() {
        System.out.println(2222);
        try {
            System.out.println("randomNum2().size():"+randomNum2().size());
            JsonArray array = randomNum2();
            for (int i = 0; i <=array.size()/1000;i++){
//                System.out.println(9999);
                timeSeries.add(new Millisecond(), array.get(i).getAsDouble());
//                System.out.println(000);
                Thread.sleep(50);
            }
        } catch (Exception e) {
        	System.out.println(e);
        }

    }

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception {
       RealTimeChart real = new RealTimeChart("","","");
       /* JFrame frame = new JFrame("Test Chart");*/
        RealTimeChart rtcp = new RealTimeChart("Random Data", "随机数", "数值");
       /* frame.getContentPane().add(rtcp, new BorderLayout().CENTER);
        frame.pack();
        frame.setVisible(true);*/
        /*frame.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent windowevent) {
                System.exit(0);
            }

        });*/
        run1();
        System.out.println(11);
        real.saveAsFile(real.getChart(),"E:\\output.jpg",100,100);
        System.out.println(22);
    }

}