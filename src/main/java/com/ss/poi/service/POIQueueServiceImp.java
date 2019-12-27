package com.ss.poi.service;

import com.ss.poi.entity.SensorPOIException;
import com.ss.poi.entity.TestResultData;

import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

/**
 * @Description
 */
public class POIQueueServiceImp {
	private static BlockingQueue<TestResultData> queue;
	private static POIQueueCustomerServiceImp customer;

	public static BlockingQueue<TestResultData> getQueue() {
		if (queue == null) {
			synchronized (POIQueueServiceImp.class) {
				if (queue == null) {
					queue = new ArrayBlockingQueue<TestResultData>(1024);
				}
			}
		}
		return queue;
	}

	public void put(TestResultData data) throws InterruptedException, SensorPOIException {
		switch (data.getFuncType()) {
		case 100030:
		case 100040:
		case 100050:
		case 100060:
		case 100070:
		case 100080:
		case 100130:
		case 100150:
		case 100180:
		case 100160:
		case 100190:
		case 100200:
		case 100210:
		case 100220:
		case 100230:
		case 100240:
		case 100260:
		case 100270:
		case 100280:
		case 100290:
		case 100530:
		case 100570:
		case 100610:
		case 101020:
		case 102020:
		case 102040:
		case 102060:
		case 102080:
		case 102090:
		case 120050:
		case 120110:
		case 120120:
		case 120130:
		case 120150:
		case 120160:
		case 120170:
		case 120030:
		case 120920:
		case 120930:
		case 120940:
		case 120950:
		case 120310:
		case 200110:
		case 300020:
		case 300030:
		case 500010:
		case 100620:
		case 100630:
		case 100640:
		case 100650:
		case 100660:
		case 100680:
		case 100710:
		case 100810:
		case 100820:
		case 500070:

			if (getQueue().size() < 500) {
				getQueue().put(data);
			} else {
				throw new SensorPOIException(-1);
			}
			break;
		}
	}

	// public TestResultData take() throws InterruptedException{
	// return this.queue.take();
	// }

	public void startDaemon() {
		if (customer == null) {
			synchronized (POIQueueServiceImp.class) {
				if (customer == null) {
					customer = new POIQueueCustomerServiceImp();
					customer.setQueue(getQueue());
					customer.start();
				}
			}
		}
	}

}
