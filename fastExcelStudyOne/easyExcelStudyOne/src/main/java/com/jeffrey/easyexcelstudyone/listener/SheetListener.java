package com.jeffrey.easyexcelstudyone.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.jeffrey.easyexcelstudyone.entity.SheetEntity;
import com.jeffrey.easyexcelstudyone.entity.StudentEntity;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/30
 * @time 10:18
 * @week 星期四
 * @description
 **/
public class SheetListener implements ReadListener<SheetEntity> {

    int i = 0;
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        System.out.println(exception.getMessage());
    }

    @Override
    public void invoke(SheetEntity SheetEntity, AnalysisContext analysisContext) {
        System.out.println("读取了第" + (++i) + "行:" + JSON.toJSONString(SheetEntity));
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("sheet1 read finished !!!");
    }
}
