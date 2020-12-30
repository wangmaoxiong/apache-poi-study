package com.wmx.poi.word;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;

/**
 * 在 Word 中创建一个简单的表格，带样式
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/30 8:46
 */
public class StyledTable {

    public static void main(String[] args) throws Exception {
        createStyledTable();
    }

    public static void createStyledTable() throws Exception {
        //XWPFDocument 是用于处理 .docx 文件的高级 API。创建一个新的空文档
        XWPFDocument xwpfDocument = new XWPFDocument();

        // 创建一个包含6行3列的新表
        int nRows = 6;
        int nCols = 3;
        XWPFTable table = xwpfDocument.createTable(nRows, nCols);

        //设置表格宽度为 100%
        table.setWidth("100%");

        // 获取表中的所有行
        List<XWPFTableRow> rows = table.getRows();
        int rowCt = 0;
        int colCt = 0;
        for (XWPFTableRow row : rows) {
            // 获取表中行的属性（trPr）
            CTTrPr trPr = row.getCtRow().addNewTrPr();
            // 设置行高
            CTHeight ht = trPr.addNewTrHeight();
            ht.setVal(BigInteger.valueOf(360));

            // 获取行中的所有列
            List<XWPFTableCell> cells = row.getTableCells();
            // 向每个单元格添加内容
            for (XWPFTableCell cell : cells) {
                // 获取表格单元格属性元素（tcPr）
                CTTcPr tcpr = cell.getCTTc().addNewTcPr();
                // 将垂直对齐设置为"居中"
                CTVerticalJc va = tcpr.addNewVAlign();
                va.setVal(STVerticalJc.CENTER);

                // 创建单元格颜色元素
                CTShd ctshd = tcpr.addNewShd();
                ctshd.setColor("auto");
                ctshd.setVal(STShd.CLEAR);
                if (rowCt == 0) {
                    //表头
                    ctshd.setFill("A7BFDE");
                } else if (rowCt % 2 == 0) {
                    // 偶数行
                    ctshd.setFill("D3DFEE");
                } else {
                    // 奇数行
                    ctshd.setFill("EDF2F8");
                }

                // 获取单元格段落列表中的第一段
                XWPFParagraph para = cell.getParagraphs().get(0);
                // 创建包含内容的文本域
                XWPFRun rh = para.createRun();
                // 根据需要设置单元格样式
                if (colCt == nCols - 1) {
                    // 最后一列的字体大小与字体
                    rh.setFontSize(10);
                    rh.setFontFamily("Courier");
                }
                if (rowCt == 0) {
                    // 表头行
                    rh.setText("header row, col " + colCt);
                    rh.setBold(true);
                    para.setAlignment(ParagraphAlignment.CENTER);
                } else {
                    // 其它行
                    rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
                }
                //局部变量 列索引加 1
                colCt++;
            }
            //新的一行开始时，列索引重置为0
            colCt = 0;
            //局部变量 行索引加 1
            rowCt++;
        }

        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "Word中创建带样式的表格" + (System.currentTimeMillis()) + ".docx");
        OutputStream os = new FileOutputStream(file);
        xwpfDocument.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
    }
}
