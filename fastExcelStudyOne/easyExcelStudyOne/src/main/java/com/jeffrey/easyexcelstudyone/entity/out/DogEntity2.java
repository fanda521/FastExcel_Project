package com.jeffrey.easyexcelstudyone.entity.out;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

import java.util.Date;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/2/18
 * @time 17:20
 * @week 星期二
 * @description
 **/
// 全局的单元格的宽度，字段也有的话，会覆盖全局的，以自己设置的为准（就近原则）
@ColumnWidth(10)
@Data
public class DogEntity2 {
    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("序号")
    private Long id;

    @ExcelIgnore
    private String color;

    private Double price;

    @ColumnWidth(30)// 单独设置这个字段的宽度，覆盖全局的
    private Date birthday;
}
