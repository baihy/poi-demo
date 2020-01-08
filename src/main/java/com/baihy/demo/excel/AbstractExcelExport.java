package com.baihy.demo.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * @projectName: poi-demo
 * @packageName: com.baihy.demo.excel
 * @description:
 * @author: huayang.bai
 * @date: 2020/1/8 17:51
 */
public abstract class AbstractExcelExport extends ExcelStyle {

    private OutputStream outputStream;

    private String title;

    private SXSSFWorkbook workbook;

    public AbstractExcelExport(String title, OutputStream outputStream) {
        this.title = title;
        this.outputStream = outputStream;
    }

    private Integer sheetSize = 100;

    public void export() {
        // 大于1000行时会把之前的行写入硬盘
        workbook = new SXSSFWorkbook(1000);
        workbook.setCompressTempFiles(true);
        int sheetNum = getDataSize() % getSheetSize() == 0 ? getDataSize() / getSheetSize() :
                getDataSize() / getSheetSize() + 1;
        for (int i = 0; i < sheetNum; i++) {
            SXSSFSheet sheet = workbook.createSheet(title + "_" + i);
            fullSheet(sheet);
        }
        try {
            workbook.write(outputStream);
            outputStream.flush();// 刷新此输出流并强制将所有缓冲的输出字节写出
            workbook.dispose();// 释放workbook所占用的所有windows资源
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void fullSheet(SXSSFSheet sheet) {
        fullTitle(sheet);
        fullHeader(sheet, 1);
        fullSheelData(sheet, 2);
    }


    /**
     * 创建execl主题
     *
     * @param sheet
     */
    private void fullTitle(SXSSFSheet sheet) {
        // 创建表头
        SXSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue(title);
        titleRow.getCell(0).setCellStyle(getTitleCellStyle(this.workbook));
        // 合并title单元格
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, getColumnsSize() - 1));
    }

    private void fullHeader(SXSSFSheet sheet, int rownum) {
        // 创建表头
        SXSSFRow headerRow = sheet.createRow(rownum);
        List<String> columnsName = getColumnsName();
        for (int i = 0; i < columnsName.size(); i++) {
            headerRow.createCell(i).setCellValue(columnsName.get(i));
            headerRow.getCell(i).setCellStyle(getHeaderCellStyle(this.workbook));
            int keyLength = columnsName.get(i).getBytes().length;
            int columnLength = keyLength > DEFAULT_COLUMN_WIDTH ? keyLength : DEFAULT_COLUMN_WIDTH;
            sheet.setColumnWidth(i, columnLength * 256);
        }
        // 冻结表头
        sheet.createFreezePane(0, rownum + 1, 0, rownum + 1);
    }


    public Integer getSheetSize() {
        return sheetSize;
    }

    public abstract int getDataSize();

    public abstract int getColumnsSize();

    public abstract List<String> getColumnsName();

    public abstract void fullSheelData(SXSSFSheet sheet, int rownum);

    public abstract List<Map<String, Object>> getRowsData(int sheetIndex);

}
