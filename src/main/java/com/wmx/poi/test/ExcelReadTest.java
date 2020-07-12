package com.wmx.poi.test;

import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/9 21:03
 */
@SuppressWarnings("all")
public class ExcelReadTest {
    /**
     * 读取 excel 文件内容，逐行逐列读取
     * 注意事项：
     * 1、以前有内容，后来设为空的单元格仍然可能被 Excel和 ApachePOI 计算为有效单元格，所以实际返回的个数可能要多，比如空字符串或者null
     *
     * @throws IOException
     */
    @Test
    public void readExcelTest() throws IOException {
        File xlsFile = new File("excel1.xls");//被读取的 excel 文件
        /**
         * WorkbookFactory 工厂的 create 方法从给定的输入流创建适当的 HSSFWorkbook 或者 XSSFWorkbook。
         * 从而避免了 excel 文件的版本兼容性问题
         */
        Workbook workbook = WorkbookFactory.create(new FileInputStream(xlsFile));
        int numberOfSheets = workbook.getNumberOfSheets();//获取工作簿中电子表格的数量
        System.out.printf("工作表个数为 %s %n", numberOfSheets);

        // 获取第一页，即第一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        /**
         * getPhysicalNumberOfRows:返回物理定义的行数（不是工作表中的行数），即实际有内容的行数，比如第1、3、5、10行有数据，则返回 4
         * getFirstRowNum:获取工作表上的第一行,以前有内容，后来被设置为空的行可能仍然被Excel和apachepoi计算为行,从0开始。
         * getLastRowNum：获取工作表上的最后一行,以前有内容，后来被设置为空的行可能仍然被Excel和apachepoi计算为行,从0开始。
         * 注意 getLastRowNum 获取的是实际有内容的最后一行的索引值，并没有+1，即这一行是有内容的，for 循环时要小于等于。
         */
        int rows = sheet.getPhysicalNumberOfRows();
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        System.out.printf("工作表中有内容的行有 %s 行，[%s,%s]%n", rows, firstRowNum, lastRowNum);
        for (int i = firstRowNum; i <= lastRowNum; i++) {
            //逐个获取有内容的每一行，如果当前行中没有任何内容，则返回 null.
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            /**
             * short getFirstCellNum():获取此行包含的第一个单元格的编号，如果行不包含任何单元格，则返回-1。
             * short getLastCellNum():获取此行中包含的最后一个单元格的索引 + 1，如果该行不包含任何单元格，则返回-1。
             * 注意：以前有内容，后来设为空的单元格仍然可能被 Excel和 ApachePOI 计算为单元格，此时也会返回，
             * 因此返回的值可能比预期的要多，因为会返回一些空字符或者 null。
             */
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();

            /**
             * 获取定义的单元格数（不是实际行中的单元格数！）,比如只有列 0、4、5有值，那么就是 3
             */
            int physicalNumberOfCells = row.getPhysicalNumberOfCells();

            System.out.printf("第 %s 行有效单元格个数为 %s ,", i + 1, physicalNumberOfCells);
            for (int j = firstCellNum; j < lastCellNum; j++) {
                /**MissingCellPolicy：用于指定空白单元格以及 null 值时的处理策略
                 * RETURN_NULL_AND_BLANK: 空返回空，null 返回 null，相当于没处理一样
                 * RETURN_BLANK_AS_NULL：空 和 null 全部返回为 null
                 * CREATE_NULL_AS_BLANK：空 和 null 全部返回为空字符串，为了后期出现空指针异常，建议使用此种策略
                 */
                Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                //设置单元格类型（空白、数值、布尔值、错误或字符串）
                cell.setCellType(CellType.STRING);
                //取值时，如果类型不一致，则会抛出异常，比如是数字，getStringCellValue
                System.out.printf("列 %s [%s]，", j + 1, cell.getStringCellValue());
            }
            System.out.printf("%n");
        }
    }
}
