package com.ss.poi.entity;



public enum EPCVoltStageEnum {
    不需要设置(-1),低压(10),静态(20),动态(30);

    private int value = 0;

    private EPCVoltStageEnum(int value) {
        this.value = value;
    }

    public static EPCVoltStageEnum valueOf(int value) {
        switch (value) {
            case 10:
                return 低压;
            case 20:
                return 静态;
            case 30:
                return 动态;
            case -1:
            default:
                return 不需要设置;
        }
    }

    public int value() {
        return this.value;
    }

}