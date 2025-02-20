package com.jeffrey.easyexcelstudyone.entity.out;


import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public  class DataClass1 {
        @ExcelProperty("Name")
        private String name;

        @ExcelProperty("Age")
        private int age;
 }