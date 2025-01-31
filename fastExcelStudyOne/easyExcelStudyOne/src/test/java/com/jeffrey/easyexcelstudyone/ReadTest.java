package com.jeffrey.easyexcelstudyone;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.jeffrey.easyexcelstudyone.entity.Person2Entity;
import com.jeffrey.easyexcelstudyone.entity.PersonEntity;
import com.jeffrey.easyexcelstudyone.entity.Sheet2Entity;
import com.jeffrey.easyexcelstudyone.entity.SheetEntity;
import com.jeffrey.easyexcelstudyone.entity.StudentEntity;
import com.jeffrey.easyexcelstudyone.listener.Person2Listener;
import com.jeffrey.easyexcelstudyone.listener.PersonListener;
import com.jeffrey.easyexcelstudyone.listener.Sheet2Listener;
import com.jeffrey.easyexcelstudyone.listener.SheetListener;
import com.jeffrey.easyexcelstudyone.listener.StudentListener;
import org.junit.jupiter.api.Test;


/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/26
 * @time 13:38
 * @week 星期日
 * @description
 **/
public class ReadTest {


    @Test
    public void read004_3(){
        /**
         * 读取多个不一样的内容的sheet
         * 监听器的doAfterAllAnalysed执行 1 次
         */
        String fileName = getClass().getClassLoader().getResource("static/sheet_2.xlsx").getFile();
        ExcelReaderBuilder excelReaderBuilder = EasyExcel.read(fileName);
        ReadSheet readSheet1 = excelReaderBuilder
                .sheet(0)
                .head(SheetEntity.class)
                .registerReadListener(new SheetListener()).build();

        ReadSheet readSheet2 = excelReaderBuilder
                .sheet(1)
                .head(Sheet2Entity.class)
                .registerReadListener(new Sheet2Listener())
                .build();
        excelReaderBuilder.build().read(readSheet1, readSheet2);

    }

    @Test
    public void read004_2(){
        /**
         * 读取多个一样的内容的sheet
         * 监听器的doAfterAllAnalysed执行 n 次
         */
        String fileName = getClass().getClassLoader().getResource("static/sheet_1.xlsx").getFile();
        ExcelReaderBuilder excelReaderBuilder = EasyExcel.read(fileName);
        ReadSheet readSheet1 = excelReaderBuilder
                .sheet(0)
                .head(SheetEntity.class)
                .registerReadListener(new SheetListener()).build();

        ReadSheet readSheet2 = excelReaderBuilder
                .sheet(1)
                .head(SheetEntity.class)
                .registerReadListener(new SheetListener())
                .build();
        excelReaderBuilder.build().read(readSheet1, readSheet2);

    }

    @Test
    public void read004(){
        /**
         * 读取多个一样的内容的sheet
         * 监听器的doAfterAllAnalysed执行 n+1 次，n是sheet的个数
         */
        String fileName = getClass().getClassLoader().getResource("static/sheet_1.xlsx").getFile();
        EasyExcel.read(fileName, SheetEntity.class, new SheetListener())
                // 在 read 方法之后， 在 sheet方法之前都是设置ReadWorkbook的参数
                // 没有调用sheet()
                .doReadAll();


    }

    @Test
    public void read003(){
        /**
         * 异常处理
         */
        String fileName = getClass().getClassLoader().getResource("static/student_1.xlsx").getFile();
        EasyExcel.read(fileName, StudentEntity.class, new StudentListener())
                // 在 read 方法之后， 在 sheet方法之前都是设置ReadWorkbook的参数
                .sheet()
                .doRead();

    }

    @Test
    public void read002(){
        /**
         * 有注解 ，index & value
         */
        String fileName = getClass().getClassLoader().getResource("static/person.xlsx").getFile();
        EasyExcel.read(fileName, Person2Entity.class, new Person2Listener())
                // 在 read 方法之后， 在 sheet方法之前都是设置ReadWorkbook的参数
                .sheet()
                .doRead();

    }

    @Test
    public void read001(){
        /**
         * 没有注解 ，默认的顺序对应起来
         */
        String fileName = getClass().getClassLoader().getResource("static/person.xlsx").getFile();
        EasyExcel.read(fileName, PersonEntity.class, new PersonListener())
                // 在 read 方法之后， 在 sheet方法之前都是设置ReadWorkbook的参数
                .sheet()
                .doRead();

    }




}
