package com.jeffrey.easyexcelstudyone.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.jeffrey.easyexcelstudyone.entity.Sheet2Entity;
import com.jeffrey.easyexcelstudyone.entity.SheetEntity;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/30
 * @time 10:18
 * @week 星期四
 * @description
 **/
public class Sheet2Listener implements ReadListener<Sheet2Entity> {

    int i = 0;
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        System.out.println(exception.getMessage());
    }

    @Override
    public void invoke(Sheet2Entity SheetEntity, AnalysisContext analysisContext) {
        System.out.println("读取了第" + (++i) + "行:" + JSON.toJSONString(SheetEntity));
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("sheet2 read finished !!!");
    }
}
