package com.jeffrey.easyexcelstudyone.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

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
public class SheetEntity {

    private Long id;

    private String name;

    private Date birthday;

    private String address;
}
