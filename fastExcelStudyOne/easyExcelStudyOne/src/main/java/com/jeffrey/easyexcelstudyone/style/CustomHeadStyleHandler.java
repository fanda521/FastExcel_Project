package com.jeffrey.easyexcelstudyone.style;

import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/2/18
 * @time 17:59
 * @week 星期二
 * @description
 **/


public class CustomHeadStyleHandler extends AbstractSheetWriteHandler {
    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        workbook.getAllNames().forEach(name -> System.out.println("Sheet name: " + name));
        Sheet sheet = writeSheetHolder.getSheet();
        String sheetName = sheet.getSheetName();
        System.out.println("Sheet name 1: " + sheetName);
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            headerRow = sheet.createRow(0);
            List<List<String>> head = writeSheetHolder.getHead();
            for (int i = 0; i < head.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(head.get(i).get(0));
            }
        }
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                System.out.println(cell.getStringCellValue());
                CellStyle cellStyle = workbook.createCellStyle(); // 创建单元格样式
                Font font = workbook.createFont();
                font.setBold(true); // 设置字体加粗
                font.setItalic(true);
                font.setColor(IndexedColors.BLUE.getIndex()); // 设置字体颜色

                cellStyle.setFont(font);
                cellStyle.setAlignment(HorizontalAlignment.CENTER); // 设置水平对齐
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 设置垂直对齐
                cell.setCellStyle(cellStyle); // 应用样式到单元格
            }
        }
    }
}

