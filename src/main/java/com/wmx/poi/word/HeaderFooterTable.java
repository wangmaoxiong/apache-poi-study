package com.wmx.poi.word;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;

/**
 * 页眉、页眉中插入表格演示
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/29 20:08
 */
public class HeaderFooterTable {
    public static void main(String[] args) throws IOException {
        //XWPFDocument 是用于处理 .docx 文件的高级 API
        XWPFDocument document = new XWPFDocument();

        //创建一个带有1行3列的页眉。
        //必须在页眉或页脚中添加段落或表格，才能将文档视为有效。
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        XWPFTable table = header.createTable(1, 3);

        // 将单元格中文本周围的填充设置为1/10英寸
        int pad = (int) (0.1 * 1440);
        table.setCellMargins(pad, pad, pad, pad);

        // 将表格宽度设置为1440点的6.5英寸
        table.setWidth((int) (6.5 * 1440));

        // 无法正确设置表格或单元格宽度，表格默认为自动调整布局，这需要固定布局
        CTTbl ctTbl = table.getCTTbl();
        CTTblPr ctTblPr = ctTbl.addNewTblPr();
        CTTblLayoutType layoutType = ctTblPr.addNewTblLayout();
        layoutType.setType(STTblLayoutType.FIXED);

        // 现在为表格设置一个网格，单元格将放入网格中每个单元格宽度为3120英寸（1440英寸）或6.5英寸的三分之一
        BigInteger w = new BigInteger("3120");
        CTTblGrid grid = ctTbl.addNewTblGrid();
        for (int i = 0; i < 3; i++) {
            CTTblGridCol gridCol = grid.addNewGridCol();
            gridCol.setW(w);
        }

        // 在单元格中添加段落
        XWPFTableRow row = table.getRow(0);
        XWPFTableCell cell = row.getCell(0);
        XWPFParagraph paragraph = cell.getParagraphArray(0);
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText("大秦");

        cell = row.getCell(1);
        paragraph = cell.getParagraphArray(0);
        xwpfRun = paragraph.createRun();
        xwpfRun.setText("大楚");

        cell = row.getCell(2);
        paragraph = cell.getParagraphArray(0);
        xwpfRun = paragraph.createRun();
        xwpfRun.setText("大齐");

        // 创建带有段落的页脚
        XWPFFooter xwpfFooter = document.createFooter(HeaderFooterType.DEFAULT);
        paragraph = xwpfFooter.createParagraph();

        xwpfRun = paragraph.createRun();
        xwpfRun.setText("我是页脚");

        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "页眉页脚插入表格.docx");
        OutputStream os = new FileOutputStream(file);
        document.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
    }
}

