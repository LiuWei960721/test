package com.ss.poi.service;

import com.ss.poi.controller.WordReport;
import com.ss.poi.entity.TestResultData;

import java.util.concurrent.BlockingQueue;

/**
 * @Description
 */
public class POIQueueCustomerServiceImp extends Thread {
	private BlockingQueue<TestResultData> queue;
	
	public void setQueue(BlockingQueue<TestResultData> queue){
		this.queue=queue;
	}
	
	@Override
	public void run() {
		// TODO Auto-generated method stub
		if(this.queue==null){
			return;
		}
		while (true) {
			try {
				TestResultData data = this.queue.take();
				System.out.println(data.getDataIndex2());
				WordReport.creatWordReport(data);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}
	}

}
