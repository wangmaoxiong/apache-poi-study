package com.wmx.poi.util;

import org.apache.poi.sl.draw.DrawTableShape;
import org.apache.poi.sl.draw.SLGraphics;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.sl.usermodel.TextShape.TextPlaceholder;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Logger;

/**
 * ppt 文件生成工具类
 * 官网在线源码地址：https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hslf/examples/ApacheconEU08.java
 * 1、因为实际开发中 ppt 的格式、大小、样式、内容等不尽相同，所以无法完全提取成工具，只能作为参考，仍然需要根据实际需求进行改写
 * 2、提醒：如果生成的 ppt 文件后续有需要使用 poi 转换成 图片，则生成 ppt 文件时，对于中文推荐使用 "宋体"，否则转换成图片时很可能会乱码
 */
public final class PPTXUtils {
    private static Logger logger = Logger.getAnonymousLogger();

    public static void main(String[] args) throws Exception {
        createPPt();
    }

    /**
     * 传教 ppt 文件
     *
     * @throws Exception
     */
    public static void createPPt() throws Exception {
        logger.info("生成 ppt 文件开始...");
        /**
         * SlideShow<?,?> ppt = new HSLFSlideShow()
         * SlideShow<?,?> pptx = new XMLSlideShow()
         * HSLFSlideShow 用于生成 .ppt 格式文件，XMLSlideShow 用于生成 .pptx 格式文件
         * 亲测 HSLFSlideShow 生成的 .ppt 文件不完整，XMLSlideShow 生成的 .pptx 文件正常
         */
        try (SlideShow<?, ?> ppt = new XMLSlideShow()) {
            /**
             * setPageSize(Dimension var1)：设置幻灯片大小
             * java.awt.Dimension 表示二维尺寸
             */
            ppt.setPageSize(new Dimension(720, 540));

            slide1(ppt);//每一个 slideX 方法都是生成一页 ppt 幻灯片(slide)
            slide2(ppt);
            slide3(ppt);
            slide4(ppt);
            slide5(ppt);
            slide6(ppt);
            slide7(ppt);
            slide8(ppt);
            String ext = ppt.getClass().getName().contains("HSLF") ? "ppt" : "pptx";

            //ppt 文件默认生成在桌面,如 C:\Users\22684\Desktop\apachecon_eu_08.pptx。文件已经存在时会自动覆盖
            File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
            File file = new File(homeDirectory, "apachecon_eu_08." + ext);
            try (FileOutputStream out = new FileOutputStream(file)) {
                ppt.write(out);
                out.flush();
            }
            logger.info("生成 ppt 文件结束：" + file.getAbsolutePath());
        }
    }

    public static void slide1(SlideShow<?, ?> ppt) throws IOException {
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片

        TextBox<?, ?> box1 = slide.createTextBox();//创建一个文本框
        box1.setTextPlaceholder(TextPlaceholder.CENTER_TITLE);//设置"中心标题占位符形状文本"
        box1.setText("梅山龙宫");
        //设置锚点：定位此形状在绘图画布中的位置，坐标以点表示。
        //java.awt.Rectangle.Rectangle(int, int, int, int)：矩形
        box1.setAnchor(new Rectangle(50, 30, 620, 100));
        box1.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily("宋体");//设置字体

        TextBox<?, ?> box2 = slide.createTextBox();//创建文本框
        box2.setTextPlaceholder(TextPlaceholder.CENTER_BODY);//设置 "中心正文占位符形状文本"
        box2.setText("梅山龙宫，是湖南最大山脉雪峰山（古称梅山）的腹地景区，是国家AAAA级旅游景区、中国国家级风景名胜区、国家自然与文化双遗产、湖南省新潇湘八景景区、首届湖南大众最爱旅游目的地。");
        box2.setAnchor(new Rectangle(55, 150, 600, 300));
        /**
         * List<P> getTextParagraphs()：返回此文本框的文本段落，有几个段落，List 中就有几个元素
         * List<T> getTextRuns()：获取此文本块中包含的文本运行
         * setFontFamily(String typeface)：设置此文本运行的字体或字体名称
         * setFontSize(Double fontSize)：在此文本运行中直接设置字体大小，如果给定为空，则字体大小默认为幻灯片布局中给定的值
         */
        box2.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily("宋体");

        TextBox<?, ?> box3 = slide.createTextBox();

        box3.getTextParagraphs().get(0).getTextRuns().get(0).setFontSize(32d);//设置字体大小
        box3.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily("宋体");//设置字体名称
        box3.setText("白溪镇宣\r2018-09-09");
        box3.setHorizontalCentered(true);//设置段落是否水平居中
        box3.setAnchor(new Rectangle(500, 450, 200, 100));
    }

    public static void slide2(SlideShow<?, ?> ppt) throws IOException {
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片
        TextBox<?, ?> box1 = slide.createTextBox();//创建文本框
        box1.setTextPlaceholder(TextPlaceholder.TITLE);//设置 "标题占位符形状文本"
        box1.setText("What is HSLF?");//设置文本
        box1.setFillColor(Color.ORANGE);//设置整个文本框填充色（注意不是字体眼色）
        box1.setAnchor(new Rectangle(30, 20, 650, 80));
        box1.getTextParagraphs().get(0).getTextRuns().get(0).setFontColor(Color.white);//设置字体颜色

        TextBox<?, ?> box2 = slide.createTextBox();
        box2.setTextPlaceholder(TextPlaceholder.BODY);//设置 "正文占位符形状文本"
        box2.setText("HorribleSLideshowFormat is the POI Project's pure Java implementation of the Powerpoint binary file format. \rPOI sub-project since 2005\rStarted by Nick Burch, Yegor Kozlov joined soon after");
        box2.setAnchor(new Rectangle(30, 120, 650, 360));

        int textParagraphsSize = box2.getTextParagraphs().size();//因为有多个段落，所以遍历设置字体
        for (int i = 0; i < textParagraphsSize; i++) {
            box2.getTextParagraphs().get(i).getTextRuns().get(0).setFontFamily("Microsoft Himalaya");//设置字体
        }
    }

    public static void slide3(SlideShow<?, ?> ppt) throws IOException {
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片
        TextBox<?, ?> box1 = slide.createTextBox();//创建文本框
        box1.setTextPlaceholder(TextPlaceholder.TITLE);//设置 "标题占位符形状文本"
        box1.setText("HSLF in a Nutshell");
        box1.setAnchor(new Rectangle(36, 15, 648, 65));//
        box1.setHorizontalCentered(true);//设置文本框内容水平居中。默认左对齐
        box1.setFillColor(Color.ORANGE);//设置文本框填充颜色（背景色）
        box1.getTextParagraphs().get(0).getTextRuns().get(0).setFontColor(Color.white);//设置字体颜色

        TextBox<?, ?> box2 = slide.createTextBox();
        box2.setTextPlaceholder(TextPlaceholder.BODY);
        box2.setText(
                "HSLF provides a way to read, create and modify MS PowerPoint presentations\r" +
                        "Pure Java API - you don't need PowerPoint to read and write *.ppt files\r" +
                        "Comprehensive support of PowerPoint objects\r" +
                        "Rich text\r" +
                        "Tables\r" +
                        "Shapes\r" +
                        "Pictures\r" +
                        "Master slides\r" +
                        "Access to low level data structures");
        List<? extends TextParagraph<?, ?, ?>> tp = box2.getTextParagraphs();
        for (int i : new byte[]{0, 1, 2, 8}) {//设置第 0，1，2，8 段落
            TextRun textRun = tp.get(i).getTextRuns().get(0);
            textRun.setFontSize(28d);//设置当前段落的字体大小
            textRun.setFontFamily("Microsoft Himalaya");//设置当前段落的字体
        }
        for (int i : new byte[]{3, 4, 5, 6, 7}) {//设置第 3, 4, 5, 6, 7 段落
            TextRun textRun = tp.get(i).getTextRuns().get(0);
            textRun.setFontSize(24d);
            textRun.setFontFamily("Microsoft Himalaya");
            /**
             * void setIndentLevel(int level)：指定此段落将遵循的特定级别文本属性，设置段落级别，级别登记有 0,1,2,3,4
             * 默认等级为等级 0，等级1就是它下面的一个子段落，以此类推
             */
            tp.get(i).setIndentLevel(1);
        }
        box2.setAnchor(new Rectangle(36, 80, 648, 400));
    }

    public static void slide4(SlideShow<?, ?> ppt) throws IOException {
        Dimension dimension = ppt.getPageSize();//获取整个PPT页面的尺寸，即整个白底页面的大小
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片
        int numRows = 4, numCols = 3;//设置表格行与列
        TableShape<?, ?> tableShape = slide.createTable(numRows, numCols);//创建一个表格
        /**设置表格的尺寸，通过简单的计算，让表格水平居中。需要根据实际情况进行调整表格的尺寸以及行高、列宽*/
        tableShape.setAnchor(new Rectangle(
                Integer.parseInt(Math.round(dimension.getWidth() / 4 / 2) + ""),
                Integer.parseInt(Math.round(dimension.getHeight() / 10 / 2) + ""),
                Integer.parseInt(Math.round(dimension.getWidth() / 4 * 3) + ""),
                Integer.parseInt(Math.round(dimension.getHeight() / 10 * 4) + "")));

        /**设置表格每个单元格的内容*/
        for (int i = 0; i < numRows; i++) {
            for (int j = 0; j < numCols; j++) {
                TableCell<?, ?> tableCell = tableShape.getCell(i, j);//获取表格单元格
                tableCell.setText("大熊山_" + i + j);//设置表格内容
                TextRun textRun = tableCell.getTextParagraphs().get(0).getTextRuns().get(0);
                textRun.setFontFamily("宋体");//设置运行文本的字体
                /**设置偶数行偶数列的，以及奇数行与奇数列的样式*/
                if ((i % 2 == 0 && j % 2 == 0) || (i % 2 == 1 && j % 2 == 1)) {
                    tableCell.setFillColor(Color.ORANGE);//背景色为橙色
                    textRun.setFontColor(Color.white);//字体为白色
                    textRun.setFontSize(22d);//字体大小为22
                    textRun.setBold(true);//字体加粗
                } else {
                    textRun.setFontColor(Color.ORANGE);
                    textRun.setFontSize(18d);
                }
                tableCell.setVerticalAlignment(VerticalAlignment.MIDDLE);//设置垂直对齐方式为居中
                /**设置表格每列的宽度为整个ppt页面宽度除以表格列数加1*/
                tableShape.setColumnWidth(j, dimension.getWidth() / (numCols + 1));
            }
            /**设置表格每行的高度为整个ppt页面高度处以10*/
            tableShape.setRowHeight(i, dimension.getHeight() / 10);
        }

        /**
         * setAllBorders(Object... args) :设置表格格式并应用于所有单元格边界
         * setOutsideBorders(Object... args)：设置表格外部边框的样式
         * setInsideBorders(Object... args)：设置表格内部边框的样式
         */
        DrawTableShape drawTableShape = new DrawTableShape(tableShape);
        drawTableShape.setAllBorders(1.0, Color.ORANGE);

        TextBox<?, ?> box1 = slide.createTextBox();//创建文本框
        box1.setHorizontalCentered(true);//设置文本框内容居中
        box1.setText("The source code is available at\rhttp://people.apache.org/~yegor/apachecon_eu08/");
        //设置文本框位置与尺寸
        box1.setAnchor(new Rectangle(
                Integer.parseInt(Math.round(dimension.getWidth() / 3) + ""),
                Integer.parseInt(Math.round(dimension.getHeight() - 100) + ""),
                Integer.parseInt(Math.round(dimension.getWidth() / 3 * 2) + ""),
                100));

        //设置文本框段落的样式
        List<? extends TextParagraph<?, ?, ?>> textParagraphs = box1.getTextParagraphs();
        TextRun textRun;
        for (int i = 0; i < textParagraphs.size(); i++) {
            textRun = box1.getTextParagraphs().get(i).getTextRuns().get(0);
            textRun.setFontSize(24d);
            textRun.setFontFamily("Microsoft Himalaya");
            textRun.setFontColor(Color.ORANGE);
        }
    }

    public static void slide5(SlideShow<?, ?> ppt) throws IOException {
        Slide<?, ?> slide = ppt.createSlide();
        TextBox<?, ?> box1 = slide.createTextBox();
        box1.setAnchor(new Rectangle(20, 20, 650, 50));
        box1.setTextPlaceholder(TextPlaceholder.TITLE);//设置 "标题占位符形状文本"
        box1.setText("HSLF in Action - 1 ata Extraction");
        box1.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily("Microsoft Himalaya");

        //为文本设置超链接，用户可以点击链接直接通过浏览器打开
        Hyperlink<?, ?> link1 = box1.getTextParagraphs().get(0).getTextRuns().get(0).createHyperlink();
        link1.linkToUrl("http://www.apache.org");//链接到 url 地址

        TextBox<?, ?> box2 = slide.createTextBox();
        box2.setAnchor(new Rectangle(20, 100, 650, 300));
        box2.setTextPlaceholder(TextPlaceholder.BODY);
        box2.setText("Text from slides and notes\r" +
                "Images\r" +
                "Shapes and their properties (type, position in the slide, color, font, etc.\r" +
                "默认段落前面有一个实心黑点，类似 html 页面的无序列表的符号)");
        List<? extends TextParagraph<?, ?, ?>> textParagraphs = box2.getTextParagraphs();
        for (int i = 0; i < textParagraphs.size(); i++) {
            textParagraphs.get(i).getTextRuns().get(0).setFontFamily("宋体");
            textParagraphs.get(i).getTextRuns().get(0).setFontSize(22d);
        }
    }

    public static void slide6(SlideShow<?, ?> ppt) throws IOException {
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片
        /**创建左侧的文本方框*/
        TextBox<?, ?> box2 = slide.createTextBox();//幻灯片中创建文本框
        box2.setAnchor(new Rectangle(66, 150, 170, 170));
        box2.setHorizontalCentered(true);//设置文本框内容水平居中
        box2.setVerticalAlignment(VerticalAlignment.MIDDLE);//垂直居中
        box2.setText("My Java Code");//设置文本内容
        box2.setFillColor(new Color(137, 12, 76));//文本框填充颜色
        box2.setStrokeStyle(0.75, Color.WHITE);//文本框边框样式
        box2.getTextParagraphs().get(0).getTextRuns().get(0).setFontColor(Color.white);

        TextBox<?, ?> box3 = slide.createTextBox();
        box3.setAnchor(new Rectangle(473, 150, 170, 170));
        box3.setHorizontalCentered(true);
        box3.setVerticalAlignment(VerticalAlignment.MIDDLE);
        box3.setText("*.PPTX File");
        box3.setFillColor(Color.ORANGE);
        box3.setStrokeStyle(0.75, Color.WHITE);
        box3.getTextParagraphs().get(0).getTextRuns().get(0).setFontColor(Color.white);

        /**AutoShape<S,P> createAutoShape():使用预定义的几何图形创建新形状并将其添加到此形状容器中
         * org.apache.poi.sl.usermodel.ShapeType：其中预设了很多常见的几何图形形状
         * ShapeType.RIGHT_ARROW：表示右箭头；ShapeType.LEFT_ARROW表示左箭头*/
        AutoShape<?, ?> box4 = slide.createAutoShape();
        box4.setAnchor(new Rectangle(253, 175, 198, 60));
        box4.setShapeType(ShapeType.RIGHT_ARROW);
        box4.setFillColor(new Color(183, 71, 42));

        AutoShape<?, ?> box5 = slide.createAutoShape();
        box5.setAnchor(new Rectangle(253, 245, 198, 60));
        box5.setShapeType(ShapeType.LEFT_ARROW);
        box5.setFillColor(new Color(183, 71, 42));
    }

    public static void slide7(SlideShow<?, ?> ppt) throws IOException {
        /**bar chart data. The first value is the bar color, the second is the width
         * 携带图表数据。偶数索引是条形图的颜色，奇数索引是是宽度*/
        Object[] dataArray = new Object[]{Color.red, 100, Color.orange, 150, Color.yellow, 75, Color.green, 200,};
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片
        GroupShape<?, ?> groupShape = slide.createGroup();//创建属于此容器的一组形状（形状组）
        //java.awt.Rectangle：创建一个矩形实例
        Rectangle rectangle = new Rectangle(Integer.parseInt(100 + ""), 50, 400, 300);
        groupShape.setAnchor(rectangle);//形状组的尺寸为矩形的位置与尺寸

        /** org.apache.poi.sl.draw.SLGraphics，java.awt.Graphics2D
         * 构建 Java 图形对象，在 PPT 绘图层中转换图形调用。
         * Graphics2D 类似于一只画笔，可以设置颜色，位置，不停的绘图
         */
        Graphics2D graphics = new SLGraphics(groupShape);

        //draw a simple bar graph(绘制简单条形图)
        int x = rectangle.x + 50, y = rectangle.y + 50;
        graphics.setFont(new Font("Arial", Font.BOLD, 10));
        for (int i = 0, idx = 1; i < dataArray.length; i += 2, idx++) {
            graphics.setColor(Color.black);//设置黑色
            int width = ((Integer) dataArray[i + 1]).intValue();
            graphics.drawString("Q" + idx, x - 20, y + 20);//此时文字绘制的是黑色的
            graphics.drawString(width + "%", x + width + 20, y + 20);//此时文字绘制的的是黑色的
            graphics.setColor((Color) dataArray[i]);//改变绘图颜色
            graphics.fill(new Rectangle(x + 10, y, width, 20));//填充颜色
            y += 60;//改变 y 值
        }
        graphics.setColor(Color.orange);
        graphics.setFont(new Font("Arial", Font.BOLD, 14));
        graphics.draw(rectangle);
        graphics.drawString("Performance（性能）", x + 70, y + 40);

    }

    public static void slide8(SlideShow<?, ?> ppt) throws IOException {
        Slide<?, ?> slide = ppt.createSlide();//创建幻灯片
        TextBox<?, ?> box1 = slide.createTextBox();//创建文本框
        box1.setTextPlaceholder(TextPlaceholder.TITLE);////设置 "标题占位符形状文本"
        box1.setText("HSLF Development Plans");
        box1.setAnchor(new Rectangle(20, 20, 648, 60));
        box1.getTextParagraphs().get(0).getTextRuns().get(0).setFontFamily("Microsoft Himalaya");

        TextBox<?, ?> box2 = slide.createTextBox();
        box2.setAnchor(new Rectangle(20, 100, 648, 450));
        box2.setTextPlaceholder(TextPlaceholder.BODY);
        box2.setText("Support for more PowerPoint functionality\r" +
                "Rendering slides into java.awt.Graphics2D\r" +
                "A way to export slides into images or other formats\r" +
                "We are also working on the XWPF for the WordprocessingML (2007+) \r" +
                "format from the OOXML specification. \r" +
                "This provides read and write support for simpler files, along with text extraction capabilities. \r" +
                "Integration with Apache FOP - Formatting Objects Processor\r" +
                "Transformation of XSL-FO into PPT\r" +
                "PPT2PDF transcoder"
        );
        //一个8个段落，设置第 0，1，7为一级段落，2，3，5，8为二级段落，4，6为三级段落
        List<? extends TextParagraph<?, ?, ?>> tp = box2.getTextParagraphs();
        for (int i : new byte[]{0, 1, 7}) {
            tp.get(i).getTextRuns().get(0).setFontSize(28d);
            tp.get(i).getTextRuns().get(0).setFontFamily("Microsoft Himalaya");
        }
        for (int i : new byte[]{2, 3, 5, 8}) {
            tp.get(i).getTextRuns().get(0).setFontSize(26d);
            tp.get(i).setIndentLevel(1);
            tp.get(i).getTextRuns().get(0).setFontFamily("Microsoft Himalaya");
        }
        for (int i : new byte[]{4, 6}) {
            tp.get(i).getTextRuns().get(0).setFontSize(24d);
            tp.get(i).setIndentLevel(2);
            tp.get(i).getTextRuns().get(0).setFontFamily("Microsoft Himalaya");
        }
    }
}
