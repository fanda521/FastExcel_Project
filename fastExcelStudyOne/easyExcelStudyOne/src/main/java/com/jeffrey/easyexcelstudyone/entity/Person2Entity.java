package com.jeffrey.easyexcelstudyone.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;

import java.util.Date;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/30
 * @time 10:16
 * @week 星期四
 * @description
 **/
@Data
public class Person2Entity {
    @ExcelProperty(index = 0)
    private Long id;

    @ExcelProperty(index = 1)
    private String name;

    @ExcelProperty(value = "生日",format = "yyyy-MM-dd")
    private Date birthday;

    @ExcelProperty(index = 3)
    private String address;

}
