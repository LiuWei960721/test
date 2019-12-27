package com.ss.poi.entity;

import com.google.gson.annotations.SerializedName;
import com.ss.poi.entity.EPCVoltStageEnum;

/**
 * @Description
 */
public class EPCStage {
	@SerializedName("FuncType")
	int FuncType;
	String Caption;
	int EPCVoltStage;
	int BusVoltType;
	int EPCLoadStage;
	int RFFreqStage;
	String ExternelStage;
	int TestCount;

	public String getVoltStageStr() {
		switch (this.EPCVoltStage) {
		case 10:
			return "低压";
		case 20:
			return "静态";
		case 30:
			return "动态";
		case -1:
		default:
			return "";
		}
	}

	public String getBusVoltStr() {
		switch (this.BusVoltType) {
		case 10:
			return "中心母线";
		case 20:
			return "低母线";
		case 30:
			return "高母线";
		case -1:
		default:
			return "";
		}
	}

	public int getFuncType() {
		return FuncType;
	}

	public void setFuncType(int funcType) {
		FuncType = funcType;
	}

	public String getCaption() {
		return Caption;
	}

	public void setCaption(String caption) {
		Caption = caption;
	}

	public int getEPCVoltStage() {
		return EPCVoltStage;
	}

	public EPCVoltStageEnum getEPCVoltStageEnum() {
		return EPCVoltStageEnum.valueOf(EPCVoltStage);
	}

	public void setEPCVoltStage(int ePCVoltStage) {
		EPCVoltStage = ePCVoltStage;
	}

	public int getBusVoltType() {
		return BusVoltType;
	}

	public void setBusVoltType(int busVoltType) {
		BusVoltType = busVoltType;
	}

	public int getEPCLoadStage() {
		return EPCLoadStage;
	}

	public void setEPCLoadStage(int ePCLoadStage) {
		EPCLoadStage = ePCLoadStage;
	}

	public int getRFFreqStage() {
		return RFFreqStage;
	}

	public void setRFFreqStage(int rFFreqStage) {
		RFFreqStage = rFFreqStage;
	}

	public String getExternelStage() {
		return ExternelStage;
	}

	public void setExternelStage(String externelStage) {
		ExternelStage = externelStage;
	}

	public int getTestCount() {
		return TestCount;
	}

	public void setTestCount(int testCount) {
		TestCount = testCount;
	}
}