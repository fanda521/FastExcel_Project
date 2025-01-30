package com.jeffrey.easyexcelstudyone;

import com.alibaba.excel.EasyExcel;
import com.jeffrey.easyexcelstudyone.entity.PersonEntity;
import com.jeffrey.easyexcelstudyone.listener.PersonListener;
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
    public void read001(){
        String fileName = getClass().getClassLoader().getResource("static/person.xlsx").getFile();
        EasyExcel.read(fileName, PersonEntity.class, new PersonListener())
                // 在 read 方法之后， 在 sheet方法之前都是设置ReadWorkbook的参数
                .sheet()
                .doRead();

    }
}
