package com.ss.poi.util;

import java.math.BigDecimal;

/**
 * @创建人: Liu
 * @创建时间: 2020-05-07 14:39
 * @描述:对double类型的小数点精确到后4位
 */

public class DoubleUtil {
    public static Double DoubleZhuanHuan(Double d){
        BigDecimal b = new BigDecimal(d);
        d = b.setScale(4, BigDecimal.ROUND_HALF_UP).doubleValue();
        return d;
    }

}
