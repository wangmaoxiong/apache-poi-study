package com.wmx.poi.word;

import org.apache.poi.xwpf.usermodel.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * 在 Word 中创建表格工具类
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/29 21:15
 */
public class WordTableUtils {
    public static void main(String[] args) throws Exception {
        //XWPFDocument 是用于处理 .docx 文件的高级 API
        XWPFDocument xwpfDocument = new XWPFDocument();

        List<Map<String, Object>> dataList = getDataList();
        String[] headerTitle = {"序号", "名称", "类型", "长度", "描述", "是否可为空"};
        createSimpleTable(xwpfDocument, headerTitle, dataList);

        //创建一个空段落，隔离前后两个表格
        XWPFParagraph paragraph = xwpfDocument.createParagraph();

        createSimpleTable(xwpfDocument, headerTitle, dataList);

        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "Word中创建简单的表格" + (System.currentTimeMillis()) + ".docx");
        OutputStream os = new FileOutputStream(file);
        xwpfDocument.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
    }

    /**
     * Word 文档中创建表格
     *
     * @param xwpfDocument ：XWPFDocument 是用于处理 .docx 文件的高级 API
     * @param headerTitle  ：表头文本内容
     * @param dataList     ：表格正文数据
     * @throws Exception
     */
    public static void createSimpleTable(XWPFDocument xwpfDocument, String[] headerTitle, List<Map<String, Object>> dataList) {

        //创建一个指定了行和列的空表
        XWPFTable table = xwpfDocument.createTable(dataList.size() + 1, dataList.get(0).size());
        //设置表格对齐方式
        table.setTableAlignment(TableRowAlign.LEFT);
        //设置表格宽度为 100%
        table.setWidth("100%");

        //设置表头
        setHeaderRow(table, headerTitle);

        int bodyRowSize = dataList.size();
        int bodyColSize = dataList.get(0).size();

        for (int i = 0; i < bodyRowSize; i++) {
            Map<String, Object> objectMap = dataList.get(i);
            for (int j = 0; j < bodyColSize; j++) {
                Object[] objects = objectMap.values().toArray();
                if (j < objects.length) {
                    table.getRow(i + 1).getCell(j).setText(objects[j].toString());
                }
            }
        }
    }

    /**
     * 设置表头
     *
     * @param table
     */
    private static void setHeaderRow(XWPFTable table, String[] headerTitle) {
        int cellIndex = 0;
        for (String title : headerTitle) {
            XWPFTableCell headCell = table.getRow(0).getCell(cellIndex++);
            headCell.setColor("BFC6D1");
            headCell.setText(title);
        }
    }

    /**
     * 设置假数据
     *
     * @return
     */
    private static List<Map<String, Object>> getDataList() {
        List<Map<String, Object>> dataList = new ArrayList<>();
        Map<String, Object> dataMap1 = new LinkedHashMap<>(8);
        dataMap1.put("column_id", "1");
        dataMap1.put("column_name", "ID");
        dataMap1.put("data_type", "NUMBER");
        dataMap1.put("data_length", 22);
        dataMap1.put("comments", "主键");
        dataMap1.put("nullable", "N");

        Map<String, Object> dataMap2 = new LinkedHashMap<>(8);
        dataMap2.put("column_id", 2);
        dataMap2.put("column_name", "TOKEN");
        dataMap2.put("data_type", "VARCHAR2");
        dataMap2.put("data_length", 255);
        dataMap2.put("comments", "token口令");
        dataMap2.put("nullable", "N");

        Map<String, Object> dataMap3 = new LinkedHashMap<>(8);
        dataMap3.put("column_id", 3);
        dataMap3.put("column_name", "BIZ_TYPE");
        dataMap3.put("data_type", "VARCHAR2");
        dataMap3.put("data_length", 63);
        dataMap3.put("comments", "此token可访问的业务类型标识");
        dataMap3.put("nullable", "N");

        Map<String, Object> dataMap4 = new LinkedHashMap<>(8);
        dataMap4.put("column_id", 4);
        dataMap4.put("column_name", "REMARK");
        dataMap4.put("data_type", "VARCHAR2");
        dataMap4.put("data_length", 255);
        dataMap4.put("comments", "备注");
        dataMap4.put("nullable", "Y");

        Map<String, Object> dataMap5 = new LinkedHashMap<>(8);
        dataMap5.put("column_id", 5);
        dataMap5.put("column_name", "UPDATE_TIME");
        dataMap5.put("data_type", "TIMESTAMP(6)");
        dataMap5.put("data_length", 11);
        dataMap5.put("comments", "更新时间");
        dataMap5.put("nullable", "N");

        dataList.add(dataMap1);
        dataList.add(dataMap2);
        dataList.add(dataMap3);
        dataList.add(dataMap4);
        dataList.add(dataMap5);

        return dataList;
    }
}

