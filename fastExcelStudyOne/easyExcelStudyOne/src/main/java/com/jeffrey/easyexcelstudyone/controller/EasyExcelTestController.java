package com.jeffrey.easyexcelstudyone.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.MapUtils;
import com.alibaba.fastjson.JSON;
import com.jeffrey.easyexcelstudyone.entity.out.DemoTableData;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/2/20
 * @time 16:21
 * @week 星期四
 * @description
 **/

@RestController
@RequestMapping("/easyExcelTest")
public class EasyExcelTestController {

    /**
     * 文件下载并且失败的时候返回json（默认失败了会返回一个有部分数据的Excel）
     *
     * @since 2.1.1
     * localhost:8060/easyExcelTest/downloadFailedUsingJson
     */
    @GetMapping("downloadFailedUsingJson")
    public void downloadFailedUsingJson(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            List<DemoTableData> demoTableDataList = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                int index = i + 1;
                DemoTableData demoStyleData = new DemoTableData("字符串" + index, new Date(), 0.56 * i);
                demoTableDataList.add(demoStyleData);
            }
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream(), DemoTableData.class).autoCloseStream(Boolean.FALSE).sheet("模板")
                    .doWrite(demoTableDataList);
            throw new RuntimeException();
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = MapUtils.newHashMap();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }

    /**
     * 文件下载（失败了会返回一个有部分数据的Excel）
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DownloadData}
     * <p>
     * 2. 设置返回的 参数
     * <p>
     * 3. 直接写，这里注意，finish的时候会自动关闭OutputStream,当然你外面再关闭流问题不大
     */
    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        List<DemoTableData> demoTableDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DemoTableData demoStyleData = new DemoTableData("字符串" + index, new Date(), 0.56 * i);
            demoTableDataList.add(demoStyleData);
        }
        EasyExcel.write(response.getOutputStream(), DemoTableData.class).sheet("模板").doWrite(demoTableDataList);
    }
}
