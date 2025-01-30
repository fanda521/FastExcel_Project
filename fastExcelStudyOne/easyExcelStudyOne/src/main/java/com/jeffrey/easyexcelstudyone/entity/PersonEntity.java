package com.jeffrey.easyexcelstudyone.entity;

import com.alibaba.excel.annotation.ExcelProperty;
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
public class PersonEntity {
    @ExcelProperty(index = 0)
    private Long id;

    @ExcelProperty(index = 1)
    private String name;

    @ExcelProperty(index = 2)
    private Date birthday;

    @ExcelProperty(index = 3)
    private String address;

}
