package com.baihy;


import com.baihy.demo.data.ExcelColumn;
import com.baihy.demo.data.ExcelData;
import com.baihy.demo.excel.DefaultExcelExport;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @projectName: dap
 * @packageName: PACKAGE_NAME
 * @description:
 * @author: huayang.bai
 * @date: 2020/1/8 15:52
 */
public class ExcelTest {


    public static void main(String[] args) throws IOException {
        List<ExcelColumn> dataColumnList = new ArrayList<>();
        dataColumnList.add(new ExcelColumn().setColumnName("name").setShowName("姓名"));
        dataColumnList.add(new ExcelColumn().setColumnName("age").setShowName("年龄"));
        dataColumnList.add(new ExcelColumn().setColumnName("birthday").setShowName("生日"));
        List<Map<String, Object>> dataRowList = new ArrayList<>();
        int count = 205;
        for (int i = 0; i < count; i++) {
            Map<String, Object> dataMap = new HashMap<>();
            dataMap.put("name", "POI-" + i);
            dataMap.put("age", new Random().nextInt(100));
            dataMap.put("birthday", new Date());
            dataRowList.add(dataMap);
        }
        ExcelData excelData = new ExcelData();
        excelData.setExcelColumns(dataColumnList);
        excelData.setExcelDataList(dataRowList);



        FileOutputStream outputStream = new FileOutputStream("D:\\ExcelExportDemo\\" + UUID.randomUUID() + ".xlsx");
        DefaultExcelExport excelExportService = new DefaultExcelExport(excelData, "sheetName", outputStream);
        System.out.println("开始导出...");
        long start = System.currentTimeMillis();
        excelExportService.export();
        System.out.println("导出完成，用时：" + (System.currentTimeMillis() - start) + "毫秒");
        outputStream.close();
    }


}
