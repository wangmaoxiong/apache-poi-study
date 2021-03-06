package com.wmx.poi.test.excel;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.util.*;

/**
 * Excel 写文件测试类
 * 官网在线示例地址：https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/usermodel/examples/
 * HSSFDataValidationHelper 等 H开头的 API 支持 .xls 格式，XSSFDataValidationHelper 等 X 开头的 API 支持 .xlsx 格式。
 */
@SuppressWarnings("all")
public class ExcelWriteTest {

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
     * 演示创建工作表（sheet）。
     * 注意：
     * 1、sheet 的名称长度不要超过 31 个字符长度，也不要包含引号、括号、冒号等特殊字符
     * 2、工作表(sheet)名称在工作簿(Workbook)中必须是唯一的
     *
     * @throws IOException
     */
    @Test
    public void createSheet() throws IOException {
        //创建工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建第1张工作表，指定工作表的名称。
        workbook.createSheet("手机采购表");
        //创建第2张工作表，指定工作表的名称
        workbook.createSheet("零部件采购表");
        //创建第3张工作表，可以在后期指定工作表的名称
        workbook.createSheet();

        //、将第一个工作表隐藏，注意：即使被隐藏，它仍然占着此索引，仍然可以对他设置和读取内容，它仅仅只是不显示而已。
        workbook.setSheetHidden(0, true);

        //为第3张工作表设置名称，索引从0开始
        workbook.setSheetName(2, "VIP客户信息");

        //获取到第2个工作表，然后创建第一个单元格，并设置内容
        //工作表必须先存在，否则异常，类似数组下标越界
        workbook.getSheetAt(1).createRow(0).createCell(0).setCellValue("蚩尤后裔");

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
     * 注意一行最多创建 256 列，否则生成的是个表格
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

    /**
     * 演示设置列宽与行高
     *
     * @throws IOException
     */
    @Test
    public void setColWidth() throws IOException {
        Workbook workbook = new XSSFWorkbook();

        //方式一：为指定的列设置统一的列框
        Sheet sheet1 = workbook.createSheet("sheet1");
        Row row1 = sheet1.createRow(0);
        for (int i = 0; i < 10; i++) {
            row1.createCell(i).setCellValue("列" + (i + 1));
            //设置指定列的宽度，通过反复测试后可以估算出，270 乘以 xx，这个 xx 基本对应 excel 文件中的实际列宽
            sheet1.setColumnWidth(i, 15 * 270);
        }

        //方式二：为整个工作表设置统一的列宽与行高
        Sheet sheet2 = workbook.createSheet("sheet2");
        Row row2 = sheet2.createRow(0);
        for (int i = 0; i < 10; i++) {
            row2.createCell(i).setCellValue("列" + (i + 1));
        }
        //设置整个工作表的默认行高，经过反复测试，20 乘以 xx，这个xx基本对应 excel 文件中的实际行高
        sheet2.setDefaultRowHeight((short) (15 * 20));
        //设置整个工作表的默认列宽，设置的值基本对应 excel 文件中的实际列宽
        sheet2.setDefaultColumnWidth((short) 15);

        //方式三：根据标题内容长度为不同的列设置不同的列宽
        Sheet sheet3 = workbook.createSheet("sheet3");
        String[] titles = {"序号", "出身日期", "在职人员来源", "进入本单位时间", "级别（技术等级、薪级）工资", "公务卡开户银行", "退休费"};
        Row row3 = sheet3.createRow(0);
        //设置每列的宽度（列宽）、根据标题的内容长度不同，设置不同的列宽。
        for (int i = 0; i < titles.length; i++) {
            row3.createCell(i).setCellValue(titles[i]);
            if (titles[i].length() <= 2) {
                sheet3.setColumnWidth(i, 10 * 270);
            } else if (titles[i].length() <= 4) {
                sheet3.setColumnWidth(i, 16 * 270);
            } else if (titles[i].length() <= 6) {
                sheet3.setColumnWidth(i, 20 * 270);
            } else {
                sheet3.setColumnWidth(i, 25 * 270);
            }
        }

        //写入到文件
        FileOutputStream fileOut = new FileOutputStream(outPath + "x");
        workbook.write(fileOut);
    }

    /**
     * 设置单元格边框样式
     *
     * @throws IOException
     */
    @Test
    public void CellStyleTest() throws IOException {
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
     * 注意一行最多创建 256 列，否则生成的是个表格
     */
    @Test
    public void completeExample1() throws IOException {
        List<String> headList = Arrays.asList("编号", "型号", "生产日期", "指导价", "数量", "经办人");
        List<List<Object>> contentData = this.contentData();

        //1）创建工作簿、创建纸张、设置纸张名称
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "华为手机采购表");

        //设置每列的宽度（列宽）
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
        contentFont.setFontHeightInPoints((short) 11);//设置字体高度
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

    /**
     * 演示为单元格设置下拉选项 方式 1
     * 当下拉选项内容长度不是太长时采用此方法。
     *
     * @throws IOException
     */
    @Test
    public void dropDowns() throws IOException {
        // 创建支持 .xls 格式的工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //为该工作簿创建一个新工作表
        HSSFSheet sheet = workbook.createSheet("new sheet");

        //创建第1行，行索引从0开始
        HSSFRow row = sheet.createRow(0);
        //创建一个单元格并在其中输入值，0 表示第1列，值不能为 null
        row.createCell(0).setCellValue("国籍");

        //下拉框的内容长度有一定的大小限制，如果超过，则抛出异常：
        //IllegalArgumentException: String literals in formulas can't be bigger than 255 characters ASCII
        List<String> dropDownLsit = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            dropDownLsit.add("中国" + i);
        }
        String[] dropDowns = dropDownLsit.toArray(new String[dropDownLsit.size()]);

        // 创建"显式列表约束"
        DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(dropDowns);

        //单元格范围地址列表由一个包含范围数目的字段和范围地址列表组成。
        //四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(0, 10, 0, 0);

        //创建数据验证单元格的实用程序类
        HSSFDataValidation dataValidation = new HSSFDataValidation(cellRangeAddressList, dvConstraint);
        //为工作簿设置数据验证对象
        sheet.addValidationData(dataValidation);

        //输出文件
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示为单元格设置下拉选项 方式 2。
     * 有效解决下拉选项内容太长的问题，思路是：不再为下拉框直接设置内容，而是将内容设置在其它地方，比如其它 sheet 中，然后通过引用的方式进行关联，
     * 这样无论下拉选项有多少内容都不会再报错。
     * 编码上与方式1基本类似，只是稍有变化。
     *
     * @throws IOException
     */
    @Test
    public void dropDowns2() throws IOException {
        // 创建支持 .xls 格式的工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //为该工作簿创建一个新工作表
        HSSFSheet sheet = workbook.createSheet("天上人间");

        //创建第1行，行索引从0开始
        HSSFRow row = sheet.createRow(0);
        //创建一个单元格并在其中输入值，0 表示第1列，值不能为 null
        row.createCell(0).setCellValue("年龄");

        //获取下拉选项的值
        String[] dropDowns = this.getDropDowns();

        // 创建一个新的工作表，这个工作表可以让它显示着，不过通常会让它隐藏，这样用户就看不到这个存储下拉选项的工作表
        // 被关联的工作表的名称最好不要有空格，否则关联时容易找不到目标数据而报错
        String hiddenSheet = "hiddenSheet";
        HSSFSheet hidden_sheet = workbook.createSheet(hiddenSheet);
        // 设置指定索引的工作表显示或者隐藏，true 是隐藏，flase 是显示
        workbook.setSheetHidden(1, true);
        // 在新的工作表中创建下拉选项的值，为了关联的时候方便，这里的值让它竖着创建内容。即第一列、第二列、...
        for (int i = 0; i < dropDowns.length; i++) {
            hidden_sheet.createRow(i).createCell(0).setCellValue(dropDowns[i]);
        }

        /**
         * 设置公式引用
         * 1、格式："工作表名称!$起始列号$起始行号:$结束列号$结束行号"
         * 2、观察 excel 表格会发现：excel 的列使用的是大写字母，如：A、B、C、D\...X、Y、Z、AA、AB、AC、...AX、AY、AZ、BA、BB、BC...以此类推
         *     excel 的行使用的阿拉伯数字，如，1、2、3、...8、9、10、11、12、...、100、101、102...、999、1000、1001、...
         * 3、现在的目的是指定某些单元格的下拉选项引用文件中其它位置的数据，如：
         *     sheet2!$A$1:$A$50 ：表示引用名称为 sheet2 的工作表中 A1到A50 之间的数据，包括A1和A50，也就是第1列前50个单元格的内容
         *     sheet3!$B$2:$B$35 ：表示引用名称为 sheet2 的工作表中 B2到B35 之间的数据，包括B2和B35，也就是第2列中第[2,35]的单元格内容
         */
        String formula = hiddenSheet + "!$A$1:$A$" + dropDowns.length;
        //用于处理数据验证的助手。HSSFDataValidationHelper 支持 .xls 格式，XSSFDataValidationHelper 支持 .xlsx 格式
        HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper(sheet);
        //创建公式列表约束
        DataValidationConstraint dataValidationConstraint = dvHelper.createFormulaListConstraint(formula);
        //单元格范围地址列表由一个包含范围数目的字段和范围地址列表组成。
        //四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(1, 50, 0, 0);
        //创建验证
        DataValidation validation_agent = dvHelper.createValidation(dataValidationConstraint, cellRangeAddressList);
        //当输入下拉选项以外的值时，是否提示错误，默认为 true
        validation_agent.setShowErrorBox(true);
        //为工作簿设置数据验证对象
        sheet.addValidationData(validation_agent);

        //输出文件
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 演示为单元格设置下拉选项 方式 3.
     * 与方式2一致，只是方式2生成的是.xls 格式文件，本方式支持 .xlsx 格式
     *
     * @throws IOException
     */
    @Test
    public void dropDowns3() throws IOException {
        // 创建支持 .xlsx 格式的工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //为该工作簿创建一个新工作表
        XSSFSheet sheet = workbook.createSheet("天上人间");

        //创建第1行，行索引从0开始
        XSSFRow row = sheet.createRow(0);
        //创建一个单元格并在其中输入值，0 表示第1列，值不能为 null
        row.createCell(0).setCellValue("年龄");

        //获取下拉选项的值
        String[] dropDowns = this.getDropDowns();

        // 创建一个新的工作表，这个工作表可以让它显示着，不过通常会让它隐藏，这样用户就看不到这个存储下拉选项的工作表
        // 被关联的工作表的名称最好不要有空格，否则关联时容易找不到目标数据而报错
        String hiddenSheet = "hiddenSheet";
        XSSFSheet hidden_sheet = workbook.createSheet(hiddenSheet);
        // 设置指定索引的工作表显示或者隐藏，true 是隐藏，flase 是显示
        workbook.setSheetHidden(1, false);
        // 在新的工作表中创建下拉选项的值，为了关联的时候方便，这里的值让它竖着创建内容。即第一列、第二列、...
        for (int i = 0; i < dropDowns.length; i++) {
            hidden_sheet.createRow(i).createCell(0).setCellValue(dropDowns[i]);
        }

        /**
         * 设置公式引用
         * 1、格式："工作表名称!$起始列号$起始行号:$结束列号$结束行号"
         * 2、观察 excel 表格会发现：excel 的列使用的是大写字母，如：A、B、C、D\...X、Y、Z、AA、AB、AC、...AX、AY、AZ、BA、BB、BC...以此类推
         *     excel 的行使用的阿拉伯数字，如，1、2、3、...8、9、10、11、12、...、100、101、102...、999、1000、1001、...
         * 3、现在的目的是指定某些单元格的下拉选项引用文件中其它位置的数据，如：
         *     sheet2!$A$1:$A$50 ：表示引用名称为 sheet2 的工作表中 A1到A50 之间的数据，包括A1和A50，也就是第1列前50个单元格的内容
         *     sheet3!$B$2:$B$35 ：表示引用名称为 sheet2 的工作表中 B2到B35 之间的数据，包括B2和B35，也就是第2列中第[2,35]的单元格内容
         */
        String formula = hiddenSheet + "!$A$1:$A$" + dropDowns.length;
        //用于处理数据验证的助手。HSSFDataValidationHelper 支持 .xls 格式，XSSFDataValidationHelper 支持 .xlsx 格式
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        //创建公式列表约束
        XSSFDataValidationConstraint dataValidationConstraint = (XSSFDataValidationConstraint) dvHelper.createFormulaListConstraint(formula);
        //单元格范围地址列表由一个包含范围数目的字段和范围地址列表组成。
        //四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(1, 50, 0, 0);
        //创建验证
        XSSFDataValidation validation_agent = (XSSFDataValidation) dvHelper.createValidation(dataValidationConstraint, cellRangeAddressList);
        //当输入下拉选项以外的值时，是否提示错误，默认为 true
        validation_agent.setShowErrorBox(false);
        //为工作簿设置数据验证对象
        sheet.addValidationData(validation_agent);

        //输出文件，文件格式设置为 .xlsx
        FileOutputStream fileOut = new FileOutputStream(outPath + "x");
        workbook.write(fileOut);
    }

    /**
     * 演示为单元格创建提示信息，当鼠标点上时，就会自动弹出提示信息，鼠标移除时，自动消失
     *
     * @throws IOException
     */
    @Test
    public void cellPrompt() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("new sheet");

        //创建第1行，行索引从0开始
        HSSFRow row = sheet.createRow(0);
        //创建一个单元格并在其中输入值，0 表示第1列，值不能为 null
        row.createCell(1).setCellValue("春运");

        // 创建"自定义公式约束"
        DVConstraint dvConstraint = DVConstraint.createCustomFormulaConstraint("BB1");

        //单元格范围地址列表由一个包含范围数目的字段和范围地址列表组成。
        //四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(1, 50, 1, 1);

        //创建数据验证单元格的实用程序类
        HSSFDataValidation dataValidation = new HSSFDataValidation(cellRangeAddressList, dvConstraint);

        dataValidation.createPromptBox("温馨提示", "道路千万条,安全第一条");
        //为工作簿设置数据验证对象
        sheet.addValidationData(dataValidation);

        //输出文件
        FileOutputStream fileOut = new FileOutputStream(outPath);
        workbook.write(fileOut);
    }

    /**
     * 获取下拉选项的值
     *
     * @return
     */
    private String[] getDropDowns() {
        //下拉框的内容长度有一定的大小限制，如果超过，则抛出异常：
        //IllegalArgumentException: String literals in formulas can't be bigger than 255 characters ASCII
        List<String> dropDownLsit = new ArrayList<>();
        for (int i = 0; i < 200; i++) {
            dropDownLsit.add((i) + "岁");
        }
        String[] dropDowns = dropDownLsit.toArray(new String[dropDownLsit.size()]);
        return dropDowns;
    }

    /**
     * 生成表格测试内容
     *
     * @return
     */
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