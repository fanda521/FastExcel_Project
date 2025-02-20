package com.jeffrey.easyexcelstudyone.utils;


import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.StringUtils;

import java.util.Objects;

/**
 * @version 1.0
 * @Aythor lucksoul 王吉慧
 * @date 2021/12/5 11:35
 * @description 获取项目路径的工具
 */
@Log4j2
public class ProjectPathUtil {

    public static void main(String[] args) {
        //Test01();
        System.out.println(getBuildTargetPath());
        System.out.println(getBuildPackagePath(ProjectPathUtil.class));
    }

    /**
     * 根据参照的类，填入参照类所在包中的文件夹路径，返回项目的绝对路径
     * @param clazz 参照类
     * @param documentPath 参照类所在包的剩余的文件路径，包括文件名及其文件格式
     * @return
     */
    public static String getFileRealPath(Class clazz,String documentPath) {
        return getBuildPackagePath(clazz) + documentPath;
    }

    /**
     * 获取编译到target的类所在的资源路径并且包括包路径
     * @return
     */
    public static String getBuildPackagePath(Class clazz) {
        return getPath(".",clazz);
    }

    /**
     * 获取项目中原来包下的类所在包的文件夹路径，不是build 在target/classes
     * @param clazz
     * @return
     */
    public static String getPackagePath(Class clazz) {
        return getBuildPackagePath(clazz).replace("target/classes/","");
    }

    /**
     * 获取编译到target的资源路径不包括包路径
     * @return
     */
    public static String getBuildTargetPath() {
        return getPath("/",ProjectPathUtil.class);
    }

    /**
     *
     * @param pathName .  /
     * @param clazz 类class
     * @return
     */
    public static String getPath(String pathName,Class clazz) {
        return  getPath(pathName,clazz,true);
    }

    /**
     *
     * @param pathName .  /
     * @param clazz 类class
     * @param flag 是否去掉路径前的 /
     * @return
     */
    public static String getPath(String pathName,Class clazz,boolean flag) {
        if (StringUtils.isBlank(pathName)) {
            log.info("传入的pathName is null !");
            return null;
        }
        if (Objects.isNull(clazz)) {
            log.info("传入的clazz is null !");
            return null;
        }
        String path = clazz.getResource(pathName).getPath();
        if (flag) {
            path = path.substring(1);
        }
        return path;
    }

    /**
     * 将类的包路径转换成反斜杠的形式 com.jeffrey.en_de.rsa3  com/jeffrey/en_de/rsa3
     * @return
     */
    public static String convertPackageNameToBackslash(String packageName) {
        if (StringUtils.isBlank(packageName)) {
            log.error("传入的参数  packageName is null !");
            return null;
        }
        return packageName.replaceAll(".","/");
    }

    /**
     * 获取指定类的包名
     * @param clazz
     * @return
     */
    public static String getPackageName(Class clazz) {
        if (Objects.isNull(clazz)) {
            log.error("传入的参数clazz is null !");
        }
        return clazz.getPackage().getName();
    }

    public static void Test01() {
//        String path = KeyPairGenUtil.class.getResource(".").getPath();
//        System.out.println("path:"+path);
//        String name = KeyPairGenUtil.class.getPackage().getName();
//        System.out.println("packageName:"+name);

        /**
         * path:/R:/Company/HuanFuTong/study/IDEA-Project/apache-stringutils-test/target/classes/com/jeffrey/en_de/rsa3/
         * packageName:com.jeffrey.en_de.rsa3
         */
    }
}
