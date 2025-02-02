package com.jeffrey.easyexcelstudyone.entity;

import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/31
 * @time 14:17
 * @week 星期五
 * @description
 **/
@Data
public class DataConvertEntity {

    private Long id;

    private String name;

    private Date birthday;

    private BigDecimal rate;
}
