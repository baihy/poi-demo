package com.baihy.demo.data;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @projectName: poi-demo
 * @packageName: com.baihy.demo.data
 * @description:
 * @author: huayang.bai
 * @date: 2020/1/8 17:47
 */

@Data
@Accessors(chain = true)
public class ExcelColumn {
    /**
     * 列名
     */
    private String columnName;
    /**
     * 列展示名
     */
    private String showName;
}
