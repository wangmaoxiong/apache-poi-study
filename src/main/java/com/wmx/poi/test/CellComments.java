package com.wmx.poi.test;

import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 演示如何使用 excel 单元格注释，Excel 注释是一种文本形状，因此插入注释与在工作表中放置文本框非常相似
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/9 15:27
 */
public class CellComments {

    public static void main(String[] args) throws IOException {
        createWorkbook(true);
        createWorkbook(false);
    }

    /**
     * 2003 的低版本使用的是 .xls 格式，使用 HSSFWorkbook 操作
     * 2007 的高版本是的是 .xlsx 格式，使用 XSSFWorkbook 操作
     * Workbook create(boolean xssf)：可以避免版本兼容问题，如果是高版本的 .xlsx ，传入 true,否则传入 false.
     *
     * @param isXls ：true 表示 .xls 格式，false 表示 .xlsx 格式
     * @throws IOException
     */
    private static void createWorkbook(boolean isXls) throws IOException {
        Workbook workbook = WorkbookFactory.create(isXls);
        String extension = isXls ? ".xls" : ".xlsx";

        Sheet sheet = workbook.createSheet("格式 " + extension);
        //创建助手：用于处理 HSSF 和 XSSF 所需的各种实例的具体类的对象
        CreationHelper creationHelper = workbook.getCreationHelper();

        // 创建绘图对象，它是所有形状（包括单元格注释）的顶级容器。
        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();

        //创建第4行第2列的单元格对象
        Cell cell1 = sheet.createRow(3).createCell(1);
        cell1.setCellValue(creationHelper.createRichTextString("Hello, World"));

        //定义工作表中注释的大小和位置
        ClientAnchor clientAnchor = creationHelper.createClientAnchor();
        clientAnchor.setCol1(4);
        clientAnchor.setRow1(2);
        clientAnchor.setCol2(6);
        clientAnchor.setRow2(5);
        Comment comment1 = drawingPatriarch.createCellComment(clientAnchor);

        // 设置注释中的文本
        comment1.setString(creationHelper.createRichTextString("We can set comments in POI"));

        //设置评论作者。
        //将鼠标移到注释单元格上时，可以在状态栏中看到它
        comment1.setAuthor("Apache Software Foundation");

        // 将注释分配给单元格的第一种方法是通过Cell.setcell注释方法
        cell1.setCellComment(comment1);

        //在第7行创建另一个单元格
        Cell cell2 = sheet.createRow(6).createCell(1);
        cell2.setCellValue(36.6);

        clientAnchor = creationHelper.createClientAnchor();
        clientAnchor.setCol1(4);
        clientAnchor.setRow1(8);
        clientAnchor.setCol2(6);
        clientAnchor.setRow2(11);
        Comment comment2 = drawingPatriarch.createCellComment(clientAnchor);
        //修改注释的背景颜色，目前仅在HSSF中可用
        if (workbook instanceof HSSFWorkbook) {
            ((HSSFComment) comment2).setFillColor(204, 236, 255);
        }
        RichTextString string = creationHelper.createRichTextString("Normal body temperature");
        //将自定义字体应用于注释中的文本
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        font.setBold(true);
        font.setColor(IndexedColors.RED.getIndex());
        string.applyFont(font);

        comment2.setString(string);
        //默认情况下，注释是隐藏的。这个总是可见的。
        comment2.setVisible(true);
        comment2.setAuthor("Bill Gates");

        //为单元格分配注释的第二种方法是隐式指定其行和列
        comment2.setRow(6);
        comment2.setColumn(1);

        FileSystemView fileSystemView = FileSystemView.getFileSystemView();
        File homeDirectory = fileSystemView.getHomeDirectory();

        String path = homeDirectory.getAbsolutePath() + File.separator + "poi_comment" + extension;
        System.out.println("输出路径：" + path);

        try (FileOutputStream out = new FileOutputStream(path)) {
            workbook.write(out);
        }
    }
}