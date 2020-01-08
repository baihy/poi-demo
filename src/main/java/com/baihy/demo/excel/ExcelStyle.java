package com.baihy.demo.excel;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @projectName: poi-demo
 * @packageName: com.baihy.demo.excel
 * @description:
 * @author: huayang.bai
 * @date: 2020/1/8 18:50
 */
public abstract class ExcelStyle {

    final int DEFAULT_COLUMN_WIDTH = 17;

    private SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd日");


    String handleValue(Object obj) {
        String cellValue = "";
        if (obj == null) {
            cellValue = "";
        } else if (obj instanceof Date) {
            cellValue = format.format(obj);
        } else if (obj instanceof Float || obj instanceof Double) {
            cellValue = new BigDecimal(obj.toString()).setScale(2, BigDecimal.ROUND_HALF_UP).toString();
        } else if (obj instanceof Boolean) {
            if ((Boolean) obj) {
                cellValue = "男";
            } else {
                cellValue = "女";
            }
        } else {
            cellValue = obj.toString();
        }
        return cellValue;
    }


    CellStyle getTitleCellStyle(SXSSFWorkbook workbook) {
        // 表头1样式
        CellStyle titleStyle = workbook.createCellStyle();
        // 水平居中
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 字体
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBold(true);
        titleStyle.setFont(titleFont);
        return titleStyle;
    }


    CellStyle getHeaderCellStyle(SXSSFWorkbook workbook) {
        // head样式
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置颜色
        headerStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
        // 前景色纯色填充
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    /**
     * 单元格样式
     *
     * @param workbook
     * @return
     */
    CellStyle getCellStyle(SXSSFWorkbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        Font cellFont = workbook.createFont();
        cellFont.setBold(false);
        cellStyle.setFont(cellFont);
        return cellStyle;
    }

}
