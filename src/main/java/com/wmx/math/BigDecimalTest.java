package com.wmx.math;

import org.junit.Test;

import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/12 9:38
 */
@SuppressWarnings("all")
public class BigDecimalTest {

    /**
     * 演示金额计算时为什么不能使用浮点数 fouat 与 double，而应该使用 BigDecimal
     * 1、比如预期应该是 0.1+0.2=0.3，9.9F*100.0F=990.00F
     */
    @Test
    public void test0() {
        double d1 = 0.1;
        double d2 = 0.2;
        System.out.println(d1 + d2);//0.30000000000000004

        float f1 = 9.9F;
        float f2 = 100.0F;
        System.out.println(f1 * f2);//989.99994

        BigDecimal d3 = BigDecimal.valueOf(d1).add(BigDecimal.valueOf(d2));
        BigDecimal multiply = BigDecimal.valueOf(f1).multiply(BigDecimal.valueOf(f2));
        BigDecimal f3 = multiply.setScale(2, RoundingMode.HALF_UP);
        System.out.println(d3);//0.3
        System.out.println(f3);//990.00
    }

    /**
     * 演示常量
     */
    @Test
    public void test1() {
        BigDecimal zero = BigDecimal.ZERO;
        BigDecimal one = BigDecimal.ONE;
        BigDecimal ten = BigDecimal.TEN;
        //0、1、10
        System.out.printf("%s、%s、%s %n", zero.longValue(), one.longValue(), ten.longValue());
    }

    /**
     * 演示构造器
     */
    @Test
    public void test2() {
        float f1 = 3.14F;
        double d1 = 3.14159;

        BigDecimal bigDecimal1 = new BigDecimal(314);
        BigDecimal bigDecimal2 = new BigDecimal(31415920L);

        /**下面是不推荐的方式：
         * 1、flout 与 double 类型不推荐使用构造器方式，因为结果是不可预测的
         * 2、flout 类型也不推荐使用 valueOf(double val) 方法，应用精度提升也会导致结果不可预测
         */
        BigDecimal bigDecimal3 = new BigDecimal(f1);
        BigDecimal bigDecimal4 = new BigDecimal(d1);
        BigDecimal bigDecimal5 = BigDecimal.valueOf(f1);

        /**下面是 flout 、double 推荐的写法
         * 1、double 类型可以使用 valueOf(double val) 方法，flout 不推荐
         * 2、最好的方式是一律转成 String 类型操作
         */
        BigDecimal bigDecimal6 = BigDecimal.valueOf(d1);
        BigDecimal bigDecimal7 = new BigDecimal(String.valueOf(f1));
        BigDecimal bigDecimal8 = new BigDecimal(String.valueOf(d1));

        System.out.println("new BigDecimal(314)=" + bigDecimal1);//314
        System.out.println("new BigDecimal(31415920L)=" + bigDecimal2);//31415920
        System.out.println("new BigDecimal(" + f1 + "F)=" + bigDecimal3);//3.1400001049041748046875
        //3.14158999999999988261834005243144929409027099609375
        System.out.println("new BigDecimal(" + d1 + ")=" + bigDecimal4);//3.140000104904175
        System.out.println("BigDecimal.valueOf(" + f1 + "F)=" + bigDecimal5);//3.14159
        System.out.println("BigDecimal.valueOf(" + d1 + ")=" + bigDecimal6);//=3.14159
        System.out.println("new BigDecimal(\"" + f1 + "F\")=" + bigDecimal7);//3.14
        System.out.println("new BigDecimal(\"" + d1 + "\")=" + bigDecimal8);//3.14159
    }

    /**
     * 演示 加减乘除运算、求余数
     */
    @Test
    public void test3() {
        BigDecimal bigDecimal = new BigDecimal("-3.14159");
        BigDecimal abs = bigDecimal.abs();//求绝对值
        BigDecimal add = bigDecimal.add(BigDecimal.valueOf(300.45));//加法
        BigDecimal subtract = bigDecimal.subtract(BigDecimal.valueOf(-1000.14159));//减法
        BigDecimal multiply = bigDecimal.multiply(new BigDecimal("-202"));//乘法
        //除法，当结果有可能为无穷无尽的小数时，必须指定舍入模式，否则会抛异常。HALF_UP是四舍五入，同时保留 4 位小数
        BigDecimal divide = bigDecimal.divide(new BigDecimal("-0.2025"), 4, RoundingMode.HALF_UP);
        BigDecimal remainder = bigDecimal.remainder(BigDecimal.valueOf(3));//求余数

        System.out.println("abs=" + abs);//3.14159
        System.out.println("add=" + add);//297.30841
        System.out.println("subtract=" + subtract);//997.00000
        System.out.println("multiply=" + multiply);//634.60118
        System.out.println("divide=" + divide);//15.51402
        System.out.println("remainder=" + remainder);//-0.14159
    }

    /**
     * 演示  求绝对值、最大/小值、相反数、判断正负数、指数运算
     */
    @Test
    public void test4() {
        BigDecimal bigDecimal = BigDecimal.valueOf(3.1415926);
        BigDecimal negate = bigDecimal.negate();//求自己的负数
        BigDecimal abs = negate.abs();//求绝对值
        int signum = negate.signum();//正数返回1、负数返回-1，0返回0
        int signum1 = abs.signum();
        BigDecimal max = bigDecimal.max(BigDecimal.TEN);//求最大值
        BigDecimal min = bigDecimal.min(BigDecimal.ZERO);//求最小值
        BigDecimal pow = BigDecimal.valueOf(2).pow(10);//指数运算

        System.out.println("negate=" + negate);//-3.1415926
        System.out.println("abs=" + abs);//3.1415926
        System.out.println(signum + "," + signum1 + "," + BigDecimal.valueOf(0.00).signum());//-1,1,0
        System.out.println("max=" + max);//10
        System.out.println("min=" + min);//0
        System.out.println(pow);//1024
    }

    /**
     * 演示 大小比较,是否相等
     */
    @Test
    public void test5() {
        System.out.println(BigDecimal.valueOf(0).compareTo(BigDecimal.ZERO));//等于返回 0
        System.out.println(BigDecimal.valueOf(10).compareTo(BigDecimal.ZERO));//大于返回 1
        System.out.println(BigDecimal.valueOf(-10).compareTo(BigDecimal.ZERO));//小于返回 -1

        System.out.println(BigDecimal.valueOf(10).equals(BigDecimal.TEN));//true
        System.out.println(BigDecimal.valueOf(10).equals(BigDecimal.ONE));//false
    }

    /**
     * 演示设置小数位数与舍入模式
     * 1、只有除法运算可以直接设置结果保留的小数位数，以及舍入模式，其它操作需要使用 setScale 进行设置
     * 2、设置小数位数与舍入模式是非常有必要的
     */
    @Test
    public void test6() {
        BigDecimal salary = new BigDecimal("12001.5612");//假设为某人的薪水
        BigDecimal tax = salary.multiply(BigDecimal.valueOf(0.1535));//假设税率为 0.1535
        System.out.println("扣税：" + tax);//1842.23964420

        //实际中并不需要这么长的精度，通常保留2位或者4位就狗了
        /**
         * HALF_UP：四舍五入
         * UP：向上舍入,比如 1.1 ——>2、5.5 ——6，-1.1 ——> -2、-2.5——>-3
         * DOWN：向下舍入，比如 1.1 ——>1、5.5 ——5，-1.1 ——> -1、-2.5——>-2
         */
        BigDecimal half_up = tax.setScale(2, RoundingMode.HALF_UP);
        BigDecimal up = tax.setScale(2, RoundingMode.UP);
        BigDecimal down = tax.setScale(2, RoundingMode.DOWN);

        System.out.println(half_up + "," + up + "," + down);//1842.24,1842.24,1842.23
    }

}
