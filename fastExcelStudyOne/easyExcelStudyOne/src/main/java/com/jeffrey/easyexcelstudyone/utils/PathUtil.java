package com.jeffrey.easyexcelstudyone.utils;

import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2023/4/7
 * @time 15:34
 * @week 星期五
 * @description 处理路径的工具
 **/

@Slf4j
public class PathUtil {

    /**
     * 获取项目的绝对路径
     * @return
     */
    public static String getProjectAbstPath() {
        String property = System.getProperty("user.dir");
        return property;
    }


    public static String createFileUnderResource(String filePath) throws IOException {
        // 获取资源目录的路径
        String buildTargetPath = ProjectPathUtil.getBuildTargetPath();
        String fileRealPath = buildTargetPath + File.separator + filePath;
        File file = new File(fileRealPath);
        if (!file.exists()) {
            file.getParentFile().mkdirs();
            file.createNewFile();
        }
        return file.getAbsolutePath();
    }
}
