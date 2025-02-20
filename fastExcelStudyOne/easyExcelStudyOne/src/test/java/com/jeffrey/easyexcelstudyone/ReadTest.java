package com.jeffrey.easyexcelstudyone;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.fastjson.JSON;
import com.jeffrey.easyexcelstudyone.entity.DataConvertEntity;
import com.jeffrey.easyexcelstudyone.entity.Person2Entity;
import com.jeffrey.easyexcelstudyone.entity.PersonEntity;
import com.jeffrey.easyexcelstudyone.entity.Sheet2Entity;
import com.jeffrey.easyexcelstudyone.entity.SheetEntity;
import com.jeffrey.easyexcelstudyone.entity.StudentEntity;
import com.jeffrey.easyexcelstudyone.listener.DataConvertListener;
import com.jeffrey.easyexcelstudyone.listener.Person2Listener;
import com.jeffrey.easyexcelstudyone.listener.PersonListener;
import com.jeffrey.easyexcelstudyone.listener.Sheet2Listener;
import com.jeffrey.easyexcelstudyone.listener.SheetListener;
import com.jeffrey.easyexcelstudyone.listener.StudentListener;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;

import java.net.URL;
import java.util.List;
import java.util.Map;


/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/26
 * @time 13:38
 * @week 星期日
 * @description
 **/
@Slf4j
public class ReadTest {



    /**
     * 同步的返回，不推荐使用，如果数据量大会把数据放到内存里面
     */
    @Test
    public void synchronousRead() {

        // 输入的Excel文件路径
        URL systemResource = ClassLoader.getSystemResource("static/student_1.xlsx");
        String inputFile = systemResource.getPath();
        log.info("inputFile:{}", inputFile);

        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 同步读取会自动finish
        List<StudentEntity> list = EasyExcel.read(inputFile).head(StudentEntity.class).sheet().doReadSync();
        for (StudentEntity data : list) {
            log.info("读取到数据:{}", JSON.toJSONString(data));
        }
        // 这里 也可以不指定class，返回一个list，然后读取第一个sheet 同步读取会自动finish
        List<Map<Integer, String>> listMap = EasyExcel.read(inputFile).sheet().doReadSync();
        for (Map<Integer, String> data : listMap) {
            // 返回每条数据的键值对 表示所在的列 和所在列的值
            log.info("读取到数据:{}", JSON.toJSONString(data));
        }
    }


    @Test
    public void read005(){
        /**
         * 读取的数据接收的时候格式化
         */
        String fileName = getClass().getClassLoader().getResource("static/date_convert.xlsx").getFile();
        ExcelReaderBuilder excelReaderBuilder = EasyExcel.read(fileName);
        ReadSheet readSheet1 = excelReaderBuilder
                .sheet(0)
                .head(DataConvertEntity.class)
                .registerReadListener(new DataConvertListener()).build();
        excelReaderBuilder.build().read(readSheet1);

    }

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
