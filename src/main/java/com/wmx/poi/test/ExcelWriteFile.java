package com.wmx.poi.test;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;

/**
 * Excel 工具类
 * 实际中仍然需要根据情况进行改写，下面以熟悉 API 为主
 * 官网在线示例地址：https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/usermodel/examples/
 */
@SuppressWarnings("all")
public class ExcelWriteFile {

    //文件输出路径
    private String outPath;

    @Before
    public void init() {
        FileSystemView fileSystemView = FileSystemView.getFileSystemView();
        File homeDirectory = fileSystemView.getHomeDirectory();
        outPath = homeDirectory.getAbsolutePath() + File.separator + System.currentTimeMillis() + ".xls";
    }

    @After
    public void after() {
        System.out.println("输出文件：" + outPath);
    }

    /**
     * 创建作表（sheet）
     *
     * @throws IOException
     */
    @Test
    public void createSheet() throws IOException {
        //创建工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表，默认为0，即第一张工作表
        workbook.createSheet("手机采购表");
        //继续创建第二张工作表
        workbook.createSheet();
        //为第二张工作表设置名称
        workbook.setSheetName(1, "VIP客户信息");
        //写入到输出流
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示表格纸张(sheet) 缩放百分比，如 80 表示缩放到 80%
     *
     * @throws IOException
     */
    @Test
    public void setSheetZoom() throws IOException {
        int scale = 80;
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();//创建工作簿
        HSSFSheet sheet1 = hssfWorkbook.createSheet("sheet1");
        sheet1.setZoom(scale);   // 设置缩放百分比
        FileOutputStream fileOutputStream = new FileOutputStream(outPath);
        hssfWorkbook.write(fileOutputStream);//将表格写入到磁盘
        fileOutputStream.flush();
        fileOutputStream.close();
    }

    /**
     * 演示单元格内容对齐方式 - Alignment
     * org.apache.poi.ss.usermodel.HorizontalAlignment：水平对齐
     * org.apache.poi.ss.usermodel.VerticalAlignment 垂直对齐
     *
     * @throws IOException
     */
    @Test
    public void alignmentCell() throws IOException {
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

        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示如何使用字体演示
     *
     * @throws IOException
     */
    @Test
    public void WorkingWithFonts() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("new sheet");

        // 创建一行并在其中放置一些单元格，行从 0 开始。
        HSSFRow row = sheet.createRow(1);

        // 创建一个新字体并修改它。
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 24);//字体大小
        font.setFontName("新宋体");//字体
        font.setItalic(true);//设置是否使用斜体，默认 false
        font.setStrikeout(true);//设置是否在文本中使用删除线水平线，默认 false

        // 字体被设置为一个样式，所以创建一个新的样式来使用。
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);

        // 创建一个单元格并在其中输入值，然后为单元格设置样式
        HSSFCell cell = row.createCell(1);
        cell.setCellValue("原件 8880 元，现价只要 98 带回家!");
        cell.setCellStyle(style);

        // 输出到文件中
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示设置行高，以及内容自动换行
     *
     * @throws IOException
     */
    @Test
    public void setHeight() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();//创建工作簿
        HSSFSheet sheet = workbook.createSheet();//创建工作表
        HSSFCellStyle cellStyle = workbook.createCellStyle();//创建样式

        // 开启内容自动换行，必须开启自动换行，下面的右斜杠换行，以及内容超出时自动换行才会生效.
        cellStyle.setWrapText(true);

        HSSFRow row1 = sheet.createRow(0);//创建第1行
        /**
         * 设置行高，-1 表示使用默认值，表格中默认行高为 13.2
         * 使用 16 进制进行设置，0x300 生成的 excel 的行高为 38.4；0x240 生成的 excel 的行高为 28.8
         */
        row1.setHeight((short) 0x300);

        HSSFCell cell3 = row1.createCell(2);//创建第1行第3列对象
        cell3.setCellValue("使用右斜杠 \n 换行，创建新的行。");//开启自动换行时，内容中可以使用右斜杠换行
        cell3.setCellStyle(cellStyle);//为单元格设置样式

        HSSFRow row3 = sheet.createRow(2);//创建第3行对象
        row3.setHeight((short) 0x240);//设置行高

        HSSFCell cell6 = row3.createCell(5);//创建第3行第6个单元格对象

        cell6.setCellValue("内容超出时，自动换行.7月4日以来，西南地区东部至长江中下游一带再度遭遇强降雨过程。");
        cell6.setCellStyle(cellStyle);//设置样式

        //设置列框，列的宽度由 sheet 设定，而不是 cell。
        sheet.setColumnWidth(2, (int) ((50 * 8) / ((double) 1 / 20)));
        sheet.setColumnWidth(5, (int) ((50 * 8) / ((double) 1 / 20)));
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);//写入到文件中
    }

    /**
     * 演示冻结窗格，这是非常常见的操作，比如移动滚动条时，表头不动
     * createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow)
     * 1、如果 colSplit 和 rowSplit 都为零，则删除现有的冻结窗格
     * 2、colSplit：水平分割位置。
     * 3、rowspilt：垂直分割位置。
     * 4、leftmostColumn：在右窗格中可见的左列。
     * 5、topprow：在底部窗格中可见的顶行。
     *
     * @throws IOException
     */
    @Test
    public void freezePane() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet1 = workbook.createSheet("first sheet");
        HSSFSheet sheet2 = workbook.createSheet("second sheet");
        HSSFSheet sheet3 = workbook.createSheet("third sheet");

        // 冻结第1行，即上下拖动时，第1行的内容不动
        sheet1.createFreezePane(0, 1, 0, 1);
        // 冻结第1列，即左右拖动时，第1列的内容不动
        sheet2.createFreezePane(1, 0, 1, 0);
        // 冻结第前2行、前2列，即滚动条滚动时，前2行、前2列的内容不动
        sheet3.createFreezePane(2, 2);

        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 读取单元格内容
     *
     * @throws IOException
     */
    @Test
    public void getRow() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("excel1.xls");
        POIFSFileSystem fileSystem = new POIFSFileSystem(fileInputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
        HSSFSheet sheet = workbook.getSheetAt(0);//获取第一页
        //获取第3行，如果第3行没有任何数据，则返回 null，反之只要第3行任意一个单元格有数据，都会返回此行对象
        HSSFRow row = sheet.getRow(2);
        if (row == null) {
            row = sheet.createRow(2);//如果第3行为 null，则创建一行
        }
        //获取第3行的第4列，如果单元格无数据，则返回 null，否则返回此单元格
        HSSFCell cell = row.getCell(3);
        if (cell == null) {
            //创建单元格
            cell = row.createCell(3);
        }
        cell.setCellValue("深圳");
        // 写入到输出文件
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示合并单元格
     *
     * @throws IOException
     */
    @Test
    public void addMergedRegion() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("new sheet1");

        //表头样式
        HSSFCellStyle headCellStyle = workbook.createCellStyle();
        headCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        HSSFRow headRow = sheet.createRow(0);//创建表头行
        //使用 16 进制进行设置，0x300 生成的 excel 的行高为 38.4；0x240 生成的 excel 的行高为 28.8
        headRow.setHeight((short) 0X250);

        HSSFCell headCell = headRow.createCell(0);//创建第1行第1列单元格对象
        headCell.setCellValue("中华人民共和国万岁！");
        headCell.setCellStyle(headCellStyle);//设置单元格样式

        /**
         * addMergedRegion(CellRangeAddress region)：添加单元格的合并区域
         * CellRangeAddress(int firstRow, int lastRow, int firstCol, int lastCol)
         * 1、创建新单元格区域。索引从零开始。
         * 2、firstRow：第一行索引，从0开始
         * 3、lastRow：最后一行的索引，必须大于等于 firstRow
         * 4、firstCol：第一列的索引，从0开始
         * 5、lastCol：最后一列的索引，必须大于等于 firstCol
         * 如下所示合并第一行前6列的单元格
         */
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示添加超链接
     *
     * @throws IOException
     */
    @Test
    public void setHyperlink() throws IOException {
        //创建工作簿，HSSFCreationHelper 创建助手用于辅助创建超链接
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCreationHelper helper = workbook.getCreationHelper();

        //超链接的单元格样式。默认情况下，超链接为蓝色并带下划线
        HSSFCellStyle hlinkStyle = workbook.createCellStyle();
        HSSFFont hlinkFont = workbook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        hlinkStyle.setFont(hlinkFont);

        HSSFCell cell;
        HSSFSheet sheet = workbook.createSheet("Hyperlinks");

        //URL 超链接
        cell = sheet.createRow(0).createCell(0);
        cell.setCellValue("URL 超链接，点击会自动跳转到浏览器打开");
        HSSFHyperlink link = helper.createHyperlink(HyperlinkType.URL);
        link.setAddress("https://wangmaoxiong.blog.csdn.net/");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkStyle);

        //链接到当前目录中的本地文件
        cell = sheet.createRow(1).createCell(0);
        cell.setCellValue("链接到当前目录中的本地文件");
        link = helper.createHyperlink(HyperlinkType.FILE);
        link.setAddress("link1.xls");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkStyle);

        //e-mail 邮件超链接
        cell = sheet.createRow(2).createCell(0);
        cell.setCellValue("e-mail 邮件超链接");
        link = helper.createHyperlink(HyperlinkType.EMAIL);
        //注意，如果subject包含空格，请确保它们是url编码的
        link.setAddress("mailto:poi@apache.org?subject=Hyperlinks");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkStyle);

        //链接到此工作簿中的某个位置
        HSSFSheet sheet2 = workbook.createSheet("Target Sheet");
        sheet2.createRow(0).createCell(0).setCellValue("Target Cell");

        cell = sheet.createRow(3).createCell(0);
        cell.setCellValue("Worksheet Link");
        link = helper.createHyperlink(HyperlinkType.DOCUMENT);
        link.setAddress("'Target Sheet'!A1");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkStyle);

        FileOutputStream out = new FileOutputStream(outPath);
        workbook.write(out);
    }

    /**
     * 演示单元格填充，比如背景颜色
     *
     * @throws IOException
     */
    @Test
    public void setFillForegroundColor() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("new sheet");

        HSSFRow row = sheet.createRow(1);
        HSSFRow row4 = sheet.createRow(3);

        // 设置浅绿色（AQUA）背景、填充模式为大斑点（BIG_SPOTS）
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        HSSFCell cell = row.createCell(1);
        cell.setCellValue("蚩尤后裔");
        cell.setCellStyle(style);

        // 继续创建单元格样式，填充前景色为 橙色，模式为 立体前景
        style = workbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(2);
        cell.setCellValue("炎黄子孙");
        cell.setCellStyle(style);

        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);

    }

    /**
     * 演示创建单元格
     *
     * @throws IOException
     */
    @Test
    public void createCell() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("new sheet2");

        //创建第1行，行索引从0开始
        HSSFRow row = sheet.createRow(0);
        //创建一个单元格并在其中输入值，0 表示第1列
        HSSFCell cell = row.createCell(0);
        cell.setCellValue(1);

        // 继续创建其它单元格，并设置值，值不能为 null
        row.createCell(1).setCellValue(3.14159);
        row.createCell(2).setCellValue("中华民族");
        row.createCell(3).setCellValue(true);
        row.createCell(4).setCellValue("");
        row.createCell(5).setCellValue("蚩尤后裔");
        //为单元格设置错误值
        row.createCell(6).setCellErrorValue(FormulaError.NUM);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    @Test
    public void test() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        HSSFRow row = sheet.createRow(1);

        // Create a cell and put a value in it.
        HSSFCell cell = row.createCell(1);
        cell.setCellValue(4);

        // Style the cell with borders all around.
        HSSFCellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        style.setBorderTop(BorderStyle.MEDIUM_DASHED);
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(outPath);
        wb.write(fileOut);
    }

    /**
     * 完整输出 excel 示例。有表头和正文
     */
    @Test
    public void completeExample1() throws IOException {
        List<String> headList = Arrays.asList("编号", "型号", "生产日期", "指导价", "数量", "经办人");
        List<List<Object>> contentData = this.contentData();

        //1）创建工作簿、创建纸张、设置纸张名称
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "华为手机采购表");

        //设置每列的宽度
        for (int i = 0; i < headList.size(); i++) {
            sheet.setColumnWidth(i, 16 * 256);
        }

        //2）表头字体
        HSSFFont headFont = workbook.createFont();
        headFont.setFontHeightInPoints((short) 12);//设置字体高度
        headFont.setFontName("宋体");
        //HSSFColorPredefined 是一个枚举，其中提供了常用的颜色
        headFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        headFont.setBold(true);//设置字体加粗，默认字体为 Arial

        //3）正文字体
        HSSFFont contentFont = workbook.createFont();
        contentFont.setFontHeightInPoints((short) 11);
        contentFont.setFontName("宋体");
        contentFont.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());

        //4）表头样式
        HSSFCellStyle headCellStyle = workbook.createCellStyle();
        //设置填充模式。org.apache.poi.ss.usermodel.FillPatternType：单元格格式的填充图案样式的枚举值
        headCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//SOLID_FOREGROUND：实心填充
        headCellStyle.setBorderBottom(BorderStyle.THICK);//设置下边框为粗线
        //设置背景色填充颜色
        headCellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        headCellStyle.setAlignment(HorizontalAlignment.CENTER);//设置表头内容水平居中显示，默认左对齐
        headCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);////设置表头内容垂直居中显示，默认下对齐
        headCellStyle.setFont(headFont);//为样式设置字体

        //5）正文样式
        HSSFCellStyle contentCellStyle = workbook.createCellStyle();
        contentCellStyle.setBorderBottom(BorderStyle.THIN);//设置下边框为细线
        contentCellStyle.setFont(contentFont);

        //6）创建表头
        HSSFRow hssfRow = sheet.createRow(0);//创建表头行，行号从0开始
        HSSFCell hssfCell;//单元格对象
        for (int i = 0; i < headList.size(); ++i) {
            hssfCell = hssfRow.createCell(i);//创建第一行的列（单元格），列号从0开始
            hssfCell.setCellStyle(headCellStyle);
            hssfCell.setCellValue(headList.get(i));
        }

        //7）创建正文内容
        for (int i = 0; i < contentData.size(); i++) {
            List<Object> rowList = contentData.get(i);
            hssfRow = sheet.createRow(i + 1);
            for (int j = 0; j < rowList.size(); j++) {
                hssfCell = hssfRow.createCell(j);
                hssfCell.setCellStyle(contentCellStyle);
                Object cellValue = rowList.get(j);
                cellValue = cellValue == null ? "" : cellValue;
                hssfCell.setCellValue(cellValue.toString());
            }
        }

        //获取输出流，写入到文件
        FileOutputStream out = new FileOutputStream(outPath);
        workbook.write(out);
        out.close();
        workbook.close();
    }

    private List<List<Object>> contentData() {
        List<List<Object>> contentList = new ArrayList<>(8);
        for (int i = 0; i < 20; i++) {
            List<Object> rowList = new ArrayList<>();
            rowList.add(i + 1);
            rowList.add("华为P3" + i);
            rowList.add(new Date());
            rowList.add(5600);
            rowList.add(new Random().nextInt(1000));
            rowList.add("张三");
            contentList.add(rowList);
        }
        return contentList;
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
    private static void createCell(HSSFWorkbook workbook, HSSFRow hssfRow, int column, String data, HorizontalAlignment align) {
        data = data == null ? "" : data;
        HSSFCell cell = hssfRow.createCell(column);
        cell.setCellValue(data);
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(align);
        cell.setCellStyle(cellStyle);
    }
}