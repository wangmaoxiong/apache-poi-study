package com.wmx.poi.word;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;

import javax.swing.filechooser.FileSystemView;

/**
 * 页眉页脚演示
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/29 19:44
 */
public class BetterHeaderFooterExample {

    public static void main(String[] args) throws IOException {
        //XWPFDocument 是用于处理 .docx 文件的高级 API
        XWPFDocument document = new XWPFDocument();
        //XWPFParagraph 是文档、表格、页眉等中的段落
        XWPFParagraph paragraph = document.createParagraph();

        //createRun 表示往段落中创建一个文本区域
        //XWPFRun 对象用一组公共属性定义一个文本区域
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText("这是页眉、页脚演示示范！");
        xwpfRun.setBold(true);

        xwpfRun = paragraph.createRun();
        xwpfRun.setText("This is a demonstration of header and footers!");

        // 创建给定类型的页眉，HeaderFooterType 枚举有3个值：DEFAULT 每页都插入，EVEN 偶数页插入，FIRST 第一页插入
        XWPFHeader head = document.createHeader(HeaderFooterType.DEFAULT);
        // 创建页眉段落
        XWPFParagraph headParagraph = head.createParagraph();
        //设置内容水平居中
        headParagraph.setAlignment(ParagraphAlignment.CENTER);
        //段落中创建文本
        headParagraph.createRun().setText("页眉");

        XWPFFooter foot = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph footParagraph = foot.createParagraph();
        footParagraph.setAlignment(ParagraphAlignment.CENTER);
        footParagraph.createRun().setText("页脚");

        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "页眉页脚.docx");
        OutputStream os = new FileOutputStream(file);
        document.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
    }
}