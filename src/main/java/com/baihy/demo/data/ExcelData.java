package com.baihy.demo.data;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;
import java.util.Map;

/**
 * @projectName: poi-demo
 * @packageName: com.baihy.demo.data
 * @description:
 * @author: huayang.bai
 * @date: 2020/1/8 17:47
 */
@Data
@Accessors(chain = true)
public class ExcelData {

    private List<ExcelColumn> excelColumns;

    private List<Map<String, Object>> excelDataList;

}
