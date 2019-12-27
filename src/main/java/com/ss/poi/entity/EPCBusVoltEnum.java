package com.ss.poi.entity;



public enum EPCBusVoltEnum {
    不需要设置(-1),中心母线(10),低母线(20),高母线(30);

    private int value = 0;

    private EPCBusVoltEnum(int value) {
        this.value = value;
    }

    public static EPCBusVoltEnum valueOf(int value) {
        switch (value) {
            case 10:
                return 中心母线;
            case 20:
                return 低母线;
            case 30:
                return 高母线;
            case -1:
            default:
                return 不需要设置;
        }
    }

    public int value() {
        return this.value;
    }

}