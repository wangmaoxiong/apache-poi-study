package com.wmx.poi.word;

import org.apache.poi.xwpf.usermodel.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * 在 Word 中创建一个简单的表格，不带样式
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/29 21:15
 */
public class SimpleTable {
    public static void main(String[] args) throws Exception {
        createSimpleTable();
    }

    public static void createSimpleTable() throws Exception {
        //XWPFDocument 是用于处理 .docx 文件的高级 API
        XWPFDocument xwpfDocument = new XWPFDocument();
        //创建一个指定了行和列的空表
        XWPFTable table = xwpfDocument.createTable(3, 4);

        //设置表格宽度为 100%
        table.setWidth("100%");

        //设置第2行第2列的单元格内容
        table.getRow(1).getCell(1).setText("简单的 Word 表格");

        // 表格单元格有一个段落列表；创建单元格时会创建一个初始段落。(0,0) 表示第一行第一列
        XWPFParagraph p1 = table.getRow(0).getCell(0).getParagraphs().get(0);

        XWPFRun xwpfRun1 = p1.createRun();
        //加粗
        xwpfRun1.setBold(true);
        //设置文本内容
        xwpfRun1.setText("编号");
        //设置是否为斜体
        xwpfRun1.setItalic(true);
        //设置字体
        xwpfRun1.setFontFamily("宋体");
        //设置下划线
        xwpfRun1.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
        //设置文本的位置，可为正，为负
        xwpfRun1.setTextPosition(5);

        //设置第3行第3列的文本
        table.getRow(2).getCell(2).setText("表格入门！");
        table.getRow(2).getCell(3).setText("POI进阶！");

        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "Word中创建简单的表格.docx");
        OutputStream os = new FileOutputStream(file);
        xwpfDocument.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
    }
}

