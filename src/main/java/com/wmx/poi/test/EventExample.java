package com.wmx.poi.test;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * HSSFListener 用于 HSSFRequest 和 HSSFEventFactory 的监听器接口，用户应该创建
 * 支持此接口的侦听器，并将其注册到 HSSFRequest
 *
 * @author wangMaoXiong
 * @version 1.0
 * @date 2020/7/9 14:13
 */
@SuppressWarnings({"java:S106", "java:S4823"})
public class EventExample implements HSSFListener {
    private SSTRecord sstrec;

    /**
     * 当读取的记录时，进入此回调反复，根据需要对它们进行处理
     *
     * @param record 在读取时发现的记录
     */
    @Override
    public void processRecord(org.apache.poi.hssf.record.Record record) {
        switch (record.getSid()) {
            // BOFRecord 可以表示工作表或工作簿的开头
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKBOOK) {
                    System.out.println("解析到工作簿（workbook）");
                } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                    System.out.println("解析到工作表（worksheet）");
                }
                break;
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("工作表（worksheet）名称: " + bsr.getSheetname());
                break;
            case RowRecord.sid:
                //整个表里面位于最后一列的数据就是所有的行的最后一列
                RowRecord rowRecord = (RowRecord) record;
                System.out.println("找到行，第一列位于： " + rowRecord.getFirstCol() + "，最后一列位于：" + rowRecord.getLastCol());
                break;
            case NumberRecord.sid:
                NumberRecord numberRecord = (NumberRecord) record;
                System.out.println("找到数值单元格，值为：" + numberRecord.getValue() + "，在第" + numberRecord.getRow() + " 行，第 " + numberRecord.getColumn() + " 列");
                break;
            // SSTRecord 存储 Excel 中使用的一组唯一字符串
            case SSTRecord.sid:
                sstrec = (SSTRecord) record;
                for (int k = 0; k < sstrec.getNumUniqueStrings(); k++) {
                    System.out.println("字符串表值：" + k + " = " + sstrec.getString(k));
                }
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                System.out.println("找到值为的字符串单元格:" + sstrec.getString(labelSSTRecord.getSSTIndex()));
                break;
            default:
                break;
        }
    }

    /**
     * 读一个 excel 文件，把读取的内容输出到控制台
     */
    public static void main(String[] args) throws IOException {
        // 使用指定的输入文件创建新的文件输入流
        FileInputStream fin = new FileInputStream("excel1.xls");
        // 创建 POIFSFileSystem 文件系统
        POIFSFileSystem poifsFileSystem = new POIFSFileSystem(fin);
        /**
         * DocumentInputStream createDocumentInputStream(final String documentName)
         * 1、在根条目的条目列表中打开文档，从 InputStream 中获取工作簿（excel部件）流
         * 2、documentName 要打开的文档的名称
         * 3、return 一个新打开的 DocumentInputStream
         */
        InputStream din = poifsFileSystem.createDocumentInputStream("Workbook");
        // 构造出 HSSFRequest 对象
        HSSFRequest hssfRequest = new HSSFRequest();
        // 添加自定义的侦听器监听读取所有记录
        hssfRequest.addListenerForAllRecords(new EventExample());
        // 创建事件工厂
        HSSFEventFactory factory = new HSSFEventFactory();
        // 根据文档输入流处理事件读取文档
        factory.processEvents(hssfRequest, din);
        System.out.println("excel 文档读取完成.");
    }
}