package com.baihy.demo.excel;

import com.baihy.demo.data.ExcelColumn;
import com.baihy.demo.data.ExcelData;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @projectName: poi-demo
 * @packageName: com.baihy.demo.excel
 * @description:
 * @author: huayang.bai
 * @date: 2020/1/8 18:14
 */
public class DefaultExcelExport extends AbstractExcelExport {

    private ExcelData excelData;

    public DefaultExcelExport(ExcelData excelData, String title, OutputStream outputStream) {
        super(title, outputStream);
        this.excelData = excelData;
    }


    @Override
    public int getDataSize() {
        return excelData.getExcelDataList().size();
    }

    @Override
    public int getColumnsSize() {
        return excelData.getExcelColumns().size();
    }

    @Override
    public List<String> getColumnsName() {
        return excelData.getExcelColumns().parallelStream().map(excelColumn -> excelColumn.getShowName()).collect(Collectors.toList());
    }

    @Override
    public List<Map<String, Object>> getRowsData(int sheetIndex) {
        int fromIndex = sheetIndex * getSheetSize();
        int toIndex = sheetIndex * getSheetSize() + getSheetSize();
        toIndex = (toIndex > getDataSize() ? getDataSize() : toIndex);
        return excelData.getExcelDataList().subList(fromIndex, toIndex);
    }

    @Override
    public void fullSheelData(SXSSFSheet sheet, int rownum) {
        int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet.getSheetName());
        List<Map<String, Object>> rowsData = getRowsData(sheetIndex);
        if (rowsData != null && rowsData.size() > 0) {
            for (int i = 0; i < rowsData.size(); i++) {
                SXSSFRow row = sheet.createRow(rownum++);
                Map<String, Object> rowData = rowsData.get(i);
                List<ExcelColumn> columns = this.excelData.getExcelColumns();
                for (int j = 0; j < columns.size(); j++) {
                    Object obj = rowData.get(columns.get(j).getColumnName());
                    row.createCell(j).setCellValue(obj.toString());
                }
            }
        }
    }


}
