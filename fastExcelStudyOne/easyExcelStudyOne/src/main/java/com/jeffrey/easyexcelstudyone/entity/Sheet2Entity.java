package com.jeffrey.easyexcelstudyone.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/30
 * @time 17:27
 * @week 星期四
 * @description
 **/
@Data
public class Sheet2Entity {

    @ExcelProperty(index = 0)
    private Long id;

    @ExcelProperty(index = 1)
    private String name;

    @ExcelProperty(index = 3)
    private Date birthday;

    @ExcelProperty(index = 4)
    private String address;

    @ExcelProperty(index = 2)
    private BigDecimal salary;
}
