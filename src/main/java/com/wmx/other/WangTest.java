package com.wmx.other;

import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/20 9:30
 */
public class WangTest {

    @Test
    public void test1() {
        Map<String, Object> map = new LinkedHashMap<>();
        map.put("id", "1000");
        System.out.println(map.get("id"));
        System.out.println(map.get("name"));
        System.out.println( StringUtils.trimToEmpty((String)map.get("name")));
    }

    @Test
    public void test2() {
        Map<String, Object> map = new LinkedHashMap<>();
        map.put("id", "1000");

        String id = (String)map.get("id");
        String name = (String)map.get("name");
        System.out.println(id);
        System.out.println(name.length());

    }

    @Test
    public void test3() {
        List<Map<String,Object>> dataLsit = new ArrayList<>();
        Map<String, Object> map1 = new LinkedHashMap<>();
        map1.put("id", "1000");

        Map<String, Object> map2 = new LinkedHashMap<>();
        map2.put("id", "2000");


        Map<String, Object> map3 = new LinkedHashMap<>();
        map3.put("id", "3000");

        dataLsit.add(map1);
        dataLsit.add(map2);
        dataLsit.add(map3);

        System.out.println(dataLsit);
        dataLsit.add(0,map2);

        System.out.println(dataLsit);
    }
}
