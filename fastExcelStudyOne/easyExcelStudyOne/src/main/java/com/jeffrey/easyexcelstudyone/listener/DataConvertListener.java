package com.jeffrey.easyexcelstudyone.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.jeffrey.easyexcelstudyone.entity.DataConvertEntity;
import com.jeffrey.easyexcelstudyone.entity.PersonEntity;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/30
 * @time 10:18
 * @week 星期四
 * @description
 **/
public class DataConvertListener implements ReadListener<DataConvertEntity> {

    int i = 0;
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        ReadListener.super.onException(exception, context);
    }

    @Override
    public void invoke(DataConvertEntity dataConvertEntity, AnalysisContext analysisContext) {
        System.out.println("读取了第" + (++i) + "行:" + JSON.toJSONString(dataConvertEntity) + "," + dataConvertEntity.getBirthday());
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("dataConvert read finished !!!");
    }
}
