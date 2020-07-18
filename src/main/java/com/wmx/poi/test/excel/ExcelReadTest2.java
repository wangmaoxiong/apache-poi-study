package com.wmx.poi.test.excel;

import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * {@link ExcelReadTest} 使用的是起始到结束的方式，本节是迭代器的方式，且判断单元格内容数据类型
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/18 16:07
 */
@SuppressWarnings("all")
public class ExcelReadTest2 {

    @Test
    public void test1() {
        FileInputStream fileInputStream = null;
        Workbook workbook = null;
        try {
            File xlsFile = new File("excel1.xls");
            fileInputStream = new FileInputStream(xlsFile);
            workbook = WorkbookFactory.create(new FileInputStream(xlsFile));
            Sheet sheet0 = workbook.getSheetAt(0);

            /**
             * Iterator<Row> rowIterator(): 返回物理行的迭代器，这意味着第三个元素可能不是第三行，比如第二行是未定义的
             * Iterator<Cell> cellIterator(): 返回物理定义的单元格的单元格迭代器
             */
            Iterator<Row> rowIterator = sheet0.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = (Cell) cellIterator.next();
                    /**
                     * int getRowIndex()：返回工作表中包含此单元格的行的行索引，从0开始
                     * int getColumnIndex()：返回此单元格的列索引，从0开始
                     * CellType getCellType()：返回单元格类型，所有类型参考 {@link CellType} 枚举
                     *
                     */
                    System.out.printf("[%s,%s]：", cell.getRowIndex(), cell.getColumnIndex());
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.printf("%s-> %s%n", "字符串（文本）单元格类型", cell.getRichStringCellValue().getString());
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                System.out.printf("%s-> %s%n", "数字单元格类型（日期）", cell.getDateCellValue());
                            } else {
                                System.out.printf("%s-> %s%n", "数字单元格类型（整数、小数）", cell.getNumericCellValue());
                            }
                            break;
                        case BOOLEAN:
                            System.out.printf("%s-> %s%n", "布尔单元格类型", cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            System.out.printf("%s-> %s%n", "公式单元格类型", cell.getCellFormula());
                            break;
                        case BLANK:
                            System.out.printf("%s-> %s%n", "空白单元格类型", "BLANK!");
                            break;
                        case ERROR:
                            System.out.printf("%s-> %s%n", "错误单元类型", "ERROR!");
                            break;
                        default:
                            System.out.println();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}