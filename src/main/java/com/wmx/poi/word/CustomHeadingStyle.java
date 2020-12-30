package com.wmx.poi.word;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;

/**
 * 自定义标题样式
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/30 19:26
 */
public class CustomHeadingStyle {
    public static void main(String[] args) throws IOException {
        writeSimpleDocxFile();
    }

    /**
     * @throws IOException
     */
    public static void writeSimpleDocxFile() throws IOException {
        XWPFDocument docxDocument = new XWPFDocument();

        // 老外自定义了一个名字，中文版的最好还是按照word给的标题名来，否则级别上可能会乱
        addCustomHeadingStyle(docxDocument, "TS1", 1);
        addCustomHeadingStyle(docxDocument, "TS2", 2);

        // 标题1
        XWPFParagraph paragraph = docxDocument.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("基础信息库系统");
        run.setBold(true);
        run.setFontFamily("宋体");
        run.setFontSize(22);
        paragraph.setStyle("TS1");

        // 标题2
        XWPFParagraph paragraph2 = docxDocument.createParagraph();
        run = paragraph2.createRun();
        run.setText("需求分析");
        run.setBold(true);
        run.setFontFamily("宋体");
        run.setFontSize(20);
        paragraph2.setStyle("TS2");

        // 正文
        XWPFParagraph paragraphX = docxDocument.createParagraph();
        XWPFRun runX = paragraphX.createRun();

        runX.setText("正文");
        // word写入到文件
        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "自定义文档标题样式.docx");
        OutputStream os = new FileOutputStream(file);
        docxDocument.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
        os.close();
    }

    /**
     * 自定义标题样式
     *
     * @param docxDocument ：目标文档
     * @param strStyleId   ：样式id，后期正文中的标题文本会通过样式id进行关联，而且样式名称会显示在文档的标题样式中
     * @param headingLevel ：样式级别
     */
    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }
}
