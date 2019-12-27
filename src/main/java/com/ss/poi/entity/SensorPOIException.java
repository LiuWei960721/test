package com.ss.poi.entity;

/**
 * @Description 
 */
public class SensorPOIException extends Exception {

	public SensorPOIException(int errType){
		_errType=errType;
	}
	
	private int _errType;
	public int getErrType(){
		return _errType;
	}
}
