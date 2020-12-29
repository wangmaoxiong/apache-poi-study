package com.wmx.poi.word;

import org.apache.poi.xwpf.usermodel.*;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * 由 POI XWPF API 创建的简单 Word 文档
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/12/29 20:19
 */
public class SimpleDocument {
    public static void main(String[] args) throws Exception {
        //XWPFDocument 是用于处理 .docx 文件的高级 API
        XWPFDocument document = new XWPFDocument();

        //文档中创建段落,第一段和第二段看起来像一个表格，其实不是，只是设置了边框线而已
        XWPFParagraph xwpfParagraph1 = document.createParagraph();
        //段落内容水平居中
        xwpfParagraph1.setAlignment(ParagraphAlignment.CENTER);

        //设置边框
        setBorder(xwpfParagraph1);

        //设置段落内容垂直居上
        xwpfParagraph1.setVerticalAlignment(TextAlignment.TOP);

        //段落中创建文本内容
        XWPFRun xwpfRun1 = xwpfParagraph1.createRun();
        xwpfRun1.setBold(true);
        xwpfRun1.setText("我是一个简单的 Word 文档！");
        xwpfRun1.setBold(true);
        xwpfRun1.setFontFamily("宋体");
        //设置下划线
        xwpfRun1.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
        //设置文本位置
        xwpfRun1.setTextPosition(100);

        //创建第二个段落，第一段和第二段看起来像一个表格，其实不是，只是设置了边框线而已
        XWPFParagraph xwpfParagraph2 = document.createParagraph();
        xwpfParagraph2.setAlignment(ParagraphAlignment.RIGHT);

        setBorder(xwpfParagraph2);

        XWPFRun xwpfRun21 = xwpfParagraph2.createRun();
        xwpfRun21.setText("我是第二段内容咯");
        //设置删除线穿过文本内容
        xwpfRun21.setStrikeThrough(true);
        xwpfRun21.setFontSize(20);

        XWPFRun xwpfRun22 = xwpfParagraph2.createRun();
        xwpfRun22.setText("and went away");
        xwpfRun22.setStrikeThrough(true);
        xwpfRun22.setFontSize(20);
        //设置下标
        xwpfRun22.setSubscript(VerticalAlign.SUPERSCRIPT);
        //创建第3段
        XWPFParagraph xwpfParagraph3 = document.createParagraph();
        //设置单词在行尾显示不全时，是否从下一行开始
        xwpfParagraph3.setWordWrapped(true);
        //设置分页符，内容从下一页开始
        xwpfParagraph3.setPageBreak(true);

        //设置内容水平对齐方式为：两端对齐（最常用的对齐方式）
        xwpfParagraph3.setAlignment(ParagraphAlignment.BOTH);
        /**
         * setSpacingBetween(double spacing, LineSpacingRule rule) 设置行间距
         * 如果rule是AUTO，那么间距spacing是以行为单位的，否则以点为单位
         */
        xwpfParagraph3.setSpacingBetween(1.5, LineSpacingRule.AUTO);

        //设置第一行缩进,通常段落第一行会向右两个字符的位置
        xwpfParagraph3.setIndentationFirstLine(600);

        XWPFRun xwpfRun3 = xwpfParagraph3.createRun();
        //设置文本位置
        xwpfRun3.setTextPosition(30);
        xwpfRun3.setText("一、持续抓好常态化疫情防控。各级各部门要坚持常态化精准防控和局部应急处置有机结合，控制传染源、切断传播途径、保护易感人群。克服麻痹思想、松劲心态，毫不放松抓好\"外防输入、内防反弹\"工作，减少\"两节\"期间人员流动和聚集，严防死守，确保不出现规模性输入和反弹。要加强大型会议活动规范化管理，从严控制，能不举行的尽量不举行，尽量精简人员聚集性活动，对确需举办的相关活动，要做好工作预案，落实防控措施。坚持\"人\"\"物\"同防，落实早发现、早报告、早隔离、早治疗防控要求，规范做好直接接触进口物品人员的个人防护、日常监测和定期核酸检测，对重点场所采取严格的环境监测和卫生措施，落实重点人群\"应检尽检\"。发挥发热门诊等\"哨点\"作用，强化\"两节\"期间医疗检验、院感控制、疫情处置等工作。加强健康教育，引导群众坚持科学佩戴口罩、保持社交距离、勤洗手等良好习惯。");
        //段落内容隔断，新内容会从下一页开始
        xwpfRun3.addBreak(BreakType.PAGE);
        xwpfRun3.setText("No more; and by a sleep to say we end The heart-ache and the thousand natural shocks That flesh is heir to, 'tis a consummation Devoutly to be wish'd. To die, to sleep; To sleep: perchance to dream: ay, there's the rub; .......");
        //是否设置斜体
        xwpfRun3.setItalic(false);

        XWPFRun xwpfRun5 = xwpfParagraph3.createRun();
        xwpfRun5.setTextPosition(-10);
        xwpfRun5.setText("原标题：如何做好元旦春节期间十项工作，湖南省“两办”重要通知来了！");
        //添加回车符
        xwpfRun5.addCarriageReturn();
        xwpfRun5.addCarriageReturn();
        xwpfRun5.setText("2021年是中国共产党成立100周年，是\"十四五\"规划开局之年，做好元旦春节期间各项工作十分重要。各级各部门要以习近平新时代中国特色社会主义思想为指导，全面贯彻党的十九大和十九届二中、三中、四中、五中全会精神，坚持以人民为中心，按照《中共中央办公厅国务院办公厅关于做好2021年元旦春节期间有关工作的通知》要求，统筹做好新冠肺炎疫情防控和节日期间各项工作，确保广大人民群众度过欢乐祥和的节日。经省委、省人民政府同意，现将有关事项通知如下");
        //内容隔断
        xwpfRun5.addBreak();
        xwpfRun5.setText("For who would bear the whips and scorns of time, The oppressor's wrong, the proud man's contumely,");

        xwpfRun5.addBreak(BreakClear.ALL);
        xwpfRun5.setText("切实保障节日市场供应。加强煤电油气运供需监测，保障人民群众温暖过冬。深入贯彻党中央、国务院关于扎实做好\"六稳\"工作、全面落实\"六保\"任务的决策部署，全面落实\"菜篮子\"市长负责制，加强组织调度，做好节日期间粮油肉蛋菜果奶等民生商品的保供稳价工作，加强粮源组织调度，确保市场原粮供应充足，确保成品油不脱销断档。");

        File homeDirectory = FileSystemView.getFileSystemView().getHomeDirectory();
        File file = new File(homeDirectory, "简单的 Word 文档.docx");
        OutputStream os = new FileOutputStream(file);
        document.write(os);
        System.out.println("文件输出成功：" + file.getAbsolutePath());
    }

    /**
     * @param paragraph
     */
    private static void setBorder(XWPFParagraph paragraph) {
        paragraph.setBorderBottom(Borders.DOUBLE);
        paragraph.setBorderTop(Borders.DOUBLE);

        paragraph.setBorderRight(Borders.DOUBLE);
        paragraph.setBorderLeft(Borders.DOUBLE);
        paragraph.setBorderBetween(Borders.SINGLE);
    }
}

