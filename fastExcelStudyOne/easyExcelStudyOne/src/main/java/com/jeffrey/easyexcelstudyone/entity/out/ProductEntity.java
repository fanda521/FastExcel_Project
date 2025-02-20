package com.jeffrey.easyexcelstudyone.entity.out;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


@Data
@NoArgsConstructor
@AllArgsConstructor
public  class ProductEntity {
        @ExcelProperty("姓名")
        private String name;

        @ExcelProperty("年龄")
        private Integer age;

        @ExcelProperty("商品名称")
        private String productName;

        @ExcelProperty("价格")
        private Integer price;
 }