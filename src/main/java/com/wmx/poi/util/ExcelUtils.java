package com.wmx.poi.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Excel 工具类
 * 实际中仍然需要根据情况进行改写，下面以熟悉 API 为主
 * 官网在线示例地址：https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/usermodel/examples/
 */

public class ExcelUtils {
    public static void main(String[] args) throws IOException {
//        setSheetZoom(100);
//        alignmentCell();
        bigData();
    }

    /**
     * 设置表格纸张(sheet) 的缩放百分比。
     *
     * @param scale 缩放百分百,如 80 表示缩放到 80%
     * @throws IOException
     */
    public static void setSheetZoom(int scale) throws IOException {
        scale = scale <= 0 ? 100 : scale;
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();//创建工作簿
        HSSFSheet sheet1 = hssfWorkbook.createSheet("sheet1");
        sheet1.setZoom(scale);   // 设置缩放百分比
        FileOutputStream fileOutputStream = new FileOutputStream("setSheetZoom.xls");
        hssfWorkbook.write(fileOutputStream);//将表格写入到磁盘
        fileOutputStream.flush();
        fileOutputStream.close();
    }

    /**
     * Alignment 对齐方式
     * org.apache.poi.ss.usermodel.HorizontalAlignment：水平对齐
     * org.apache.poi.ss.usermodel.VerticalAlignment 垂直对齐
     *
     * @throws IOException
     */
    @SuppressWarnings("all")
    public static void alignmentCell() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet1");
        //HSSFRow createRow(int rownum)：从纸张中创建行，rownum 从0开始
        HSSFRow row = sheet.createRow(1);//即使用 sheet 纸张中的第2行
        createCell(workbook, row, 0, "中国", HorizontalAlignment.CENTER);//居中
        createCell(workbook, row, 1, "USA", HorizontalAlignment.CENTER_SELECTION);//合并居中
        createCell(workbook, row, 2, "大秦", HorizontalAlignment.FILL);//填充对齐
        createCell(workbook, row, 3, "大汉", HorizontalAlignment.GENERAL);//常规对齐，文本数据靠左
        createCell(workbook, row, 4, "大明", HorizontalAlignment.JUSTIFY);//两端对齐
        createCell(workbook, row, 5, "大清", HorizontalAlignment.LEFT);//左对齐
        createCell(workbook, row, 6, "大唐", HorizontalAlignment.RIGHT);//右对齐

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        workbook.write(fileOut);
    }

    /**
     * 创建单元格
     *
     * @param workbook 工作簿
     * @param hssfRow  hssf 行对象
     * @param column   行中的列索引，从 0 开始
     * @param data     单元格中的内容，可以为空
     * @param align    对齐方式
     */
    @SuppressWarnings("all")
    private static void createCell(HSSFWorkbook workbook, HSSFRow hssfRow, int column, String data, HorizontalAlignment align) {
        data = data == null ? "" : data;
        HSSFCell cell = hssfRow.createCell(column);
        cell.setCellValue(data);
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(align);
        cell.setCellStyle(cellStyle);
    }

    @SuppressWarnings("all")
    public static void bigData() throws IOException {
        int rowNum = 30;//将要生成的总数行
        int cellNum = 10;//将要生成的总数列
        HSSFWorkbook workbook = new HSSFWorkbook();//创建工作簿
        HSSFSheet sheet = workbook.createSheet();//创建纸张
        workbook.setSheetName(0, "wang_sheet");//设置纸张名称

        //创建三个单元格样式
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();
        HSSFCellStyle cellStyle2 = workbook.createCellStyle();
        HSSFCellStyle cellStyle3 = workbook.createCellStyle();
        //创建两个字体
        HSSFFont font1 = workbook.createFont();
        HSSFFont font2 = workbook.createFont();

        font1.setFontHeightInPoints((short) 10);//设置字体高度
        font1.setColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());//HSSFColorPredefined 是一个枚举，其中提供了常用的颜色
        font1.setBold(true);//设置字体加粗，默认字体为 Arial

        font2.setFontHeightInPoints((short) 10);
        font2.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        font2.setBold(true);

        cellStyle1.setFont(font1);//为样式添加字体
        //设置数据格式（必须是有效格式）.将"文本"转换为Excel的格式字符串以表示文本
        cellStyle1.setDataFormat(HSSFDataFormat.getBuiltinFormat("($#,##0_);[Red]($#,##0)"));

        //设置单元格下边框的边框类型.org.apache.poi.ss.usermodel.BorderStyle 是一个枚举，其中有预定义了很多边框类型
        cellStyle2.setBorderBottom(BorderStyle.THIN);//设置下边框为细线
        //设置填充模式。org.apache.poi.ss.usermodel.FillPatternType：单元格格式的填充图案样式的枚举值
        cellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);//SOLID_FOREGROUND：实心填充
        //设置背景色填充颜色
        cellStyle2.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
        cellStyle2.setFont(font2);//为样式添加字体

        HSSFRow hssfRow;//表格行对象
        HSSFCell hssfCell;//单元格对象
        for (int i = 0; i < rowNum; i++) {
            hssfRow = sheet.createRow(i);//创建表格行
            if ((i % 2) == 0) {//偶数行
                hssfRow.setHeight((short) 0x249);//设置行高
            }
            for (int j = 0; j < cellNum; j += 2) {
                hssfCell = hssfRow.createCell(j);//创建单元格
                hssfCell.setCellValue(Math.random() * 100000000);//设置单元格值
                if (i % 2 == 0) {//偶数行
                    hssfCell.setCellStyle(cellStyle1);
                }
                hssfCell = hssfRow.createCell(j + 1);
                hssfCell.setCellValue("薪水");
                //setColumnWidth(int columnIndex, int width)：设置纸张的列宽度
                //columnIndex：要设置的列，从0开始；
                //width：以字符宽度的1/256为单位,则 256 表示一个字符宽度。Excel中的最大列宽为255个字符，长度就是 255*256
                sheet.setColumnWidth(j + 1, 256 * 10);//10个字符宽度
                if ((i % 2) == 0) {//偶数行
                    hssfCell.setCellStyle(cellStyle2);//设置单元格样式
                }
            }
        }
        //往下两行的位置画一条黑线
        rowNum++;
        rowNum++;
        hssfRow = sheet.createRow(rowNum);
        cellStyle3.setBorderBottom(BorderStyle.THICK);
        for (int i = 0; i < cellNum; i++) {
            hssfCell = hssfRow.createCell(i);
            hssfCell.setCellStyle(cellStyle3);
        }
        //获取输出流，写入到文件
        FileOutputStream out = new FileOutputStream("workbook.xls");
        workbook.write(out);
    }
}