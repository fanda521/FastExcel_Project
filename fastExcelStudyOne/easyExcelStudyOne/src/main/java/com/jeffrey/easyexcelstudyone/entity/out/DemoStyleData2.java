package com.jeffrey.easyexcelstudyone.entity.out;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.enums.poi.FillPatternTypeEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class DemoStyleData2 {
    @ExcelProperty("字符串标题")
    private String string;
    @ExcelProperty("日期标题")
    @ColumnWidth(40)
    private Date date;
    @ExcelProperty("数字标题")
    private Double doubleData;
}