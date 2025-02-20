package com.jeffrey.easyexcelstudyone.entity.out;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public  class DataClass2 {
        @ExcelProperty("Product Name")
        private String productName;

        @ExcelProperty("Price")
        private int price;
}