package com.jeffrey.easyexcelstudyone.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.fastjson.JSON;
import com.jeffrey.easyexcelstudyone.entity.PosmWmsBinImportEntity;

import java.util.Map;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/1/26
 * @time 13:57
 * @week 星期日
 * @description
 **/
public class PosmWmsBinListener implements ReadListener<PosmWmsBinImportEntity> {
    int i = 0;
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        ReadListener.super.onException(exception, context);
    }

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        ReadListener.super.invokeHead(headMap, context);
    }

    @Override
    public void invoke(PosmWmsBinImportEntity posmWmsBinImportEntity, AnalysisContext analysisContext) {
        System.out.println("读取了第" + (++i) + "行:" + JSON.toJSONString(posmWmsBinImportEntity));
    }

    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        ReadListener.super.extra(extra, context);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("finished read!!!");
    }

    @Override
    public boolean hasNext(AnalysisContext context) {
        return ReadListener.super.hasNext(context);
    }
}
