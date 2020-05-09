package com.ss.poi.test;

import java.math.BigDecimal;

/**
 * @创建人: Liu
 * @创建时间: 2020-05-07 14:22
 * @描述:
 */

public class testDouble {
    public static void main(String[] args) {
        double d = 0.37939999999999996;
        BigDecimal b = new BigDecimal(d);
        d = b.setScale(4, BigDecimal.ROUND_HALF_UP).doubleValue();
        System.out.println(d);

    }
}
