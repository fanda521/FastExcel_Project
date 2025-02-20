package com.jeffrey.easyexcelstudyone.entity.out;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

@Data
@EqualsAndHashCode
public class ComplexHeadData {
    @ExcelProperty({"主标题","数字二级标题", "字符串标题"})
    private String string;
    @ColumnWidth(20)
    @ExcelProperty({"主标题", "第二级标题", "日期标题"})
    private Date date;
    @ExcelProperty({"主标题", "第二级标题", "数字标题"})
    private Double doubleData;
}