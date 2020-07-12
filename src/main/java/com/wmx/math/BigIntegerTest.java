package com.wmx.math;

import org.junit.Test;

import java.math.BigInteger;
import java.util.Random;

/**
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/12 21:39
 */
@SuppressWarnings("all")
public class BigIntegerTest {
    /**
     * 演示常量
     */
    @Test
    public void test1() {
        BigInteger one = BigInteger.ONE;
        BigInteger zero = BigInteger.ZERO;
        BigInteger ten = BigInteger.TEN;
        System.out.printf("%s，%s，%s", one, zero, ten);//1，0，10
    }

    /**
     * 演示构造器用法
     */
    @Test
    public void test2() {
        //字符串不能包含任何其他字符（例如，空格，小数点）
        BigInteger bigInteger1 = new BigInteger("39998");
        BigInteger bigInteger2 = new BigInteger("-39998");
        System.out.println(bigInteger1 + "," + bigInteger2);//39998,-39998

        //随机生成 [0 到 2的10次方减1] 内的整数值，
        BigInteger bigInteger = new BigInteger(10, new Random());
        for (int i = 0; i < 100; i++) {
            System.out.println(new BigInteger(8, 20, new Random()));
        }
    }

    /**
     * 演示 加减乘除算数运算、求绝对值、取模运算
     */
    @Test
    public void test3() {
        BigInteger bigInteger = new BigInteger("250");
        System.out.println(bigInteger.add(BigInteger.valueOf(250)));//加法 500
        System.out.println(bigInteger.subtract(BigInteger.valueOf(150)));///减法 100
        System.out.println(bigInteger.multiply(BigInteger.valueOf(4)));//乘法 1000
        System.out.println(bigInteger.divide(BigInteger.valueOf(3)));//除法 83
        System.out.println(BigInteger.valueOf(-100).abs());//取绝对值 100
        System.out.println(BigInteger.valueOf(-10).remainder(BigInteger.valueOf(3)));//取模 -1
        System.out.println(BigInteger.valueOf(-10).mod(BigInteger.valueOf(3)));//取模 2
    }

    /**
     * 演示 大小比较，求最大值、最小值、数值是否相等
     */
    @Test
    public void test4() {
        System.out.println(BigInteger.valueOf(100).compareTo(BigInteger.ZERO));//大于：1
        System.out.println(BigInteger.valueOf(0).compareTo(BigInteger.ZERO));//等于：0
        System.out.println(BigInteger.valueOf(-100).compareTo(BigInteger.ZERO));//小于：-1

        System.out.println(BigInteger.valueOf(100).max(BigInteger.TEN));//100
        System.out.println(BigInteger.valueOf(100).min(BigInteger.TEN));//10

        System.out.println(BigInteger.valueOf(10).equals(BigInteger.TEN));//true
        System.out.println(BigInteger.valueOf(-10).equals(BigInteger.TEN));//false
    }

    /**
     * 演示 指数运算、判断正负数、生成素数、判断素数
     */
    @Test
    public void test5() {
        BigInteger bigInteger = BigInteger.valueOf(2);
        System.out.println(bigInteger.pow(10));//指数运算，1024
        System.out.println(bigInteger.negate());//求自己的负数，-2
        System.out.println(bigInteger.signum());//正数返回1，0返回0，负数返回-1

        System.out.println(BigInteger.valueOf(1000).nextProbablePrime());//1009
        System.out.println(BigInteger.valueOf(11).isProbablePrime(100));//true
        System.out.println(BigInteger.valueOf(8).isProbablePrime(100));//false

        for (int i = 0; i < 10; i++) {
            //随机生成素数
            System.out.println(i + "：" + BigInteger.probablePrime(8, new Random()));
        }
    }
}
