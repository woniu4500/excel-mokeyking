package com.li.excel;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.li.excel.utils.ExcelExportUtil;
import com.li.excel.utils.ExcelImportUtil;

/**
 * 
 * @Title: ExcelUtil.java 
 * @Package com.li.excel 
 * @Description: 外部访问工具类
 * @author leevan
 * @date 2018年11月14日 下午3:15:19
 * @version 1.0.0
 */
public class ExcelUtil {

    /**
     * 导出单sheet的excel文件
     * @param filePath
     * @param collection
     * @param clazz
     * @param tag
     */
    public static void exportToFile(String filePath, Collection<?> collection, Class<?> clazz, int tag) {
    	ExcelExportUtil.exportSingleSheetToFile(filePath,collection,tag);
    } 

    /**
     * 导出多个sheet的excel文件
     * @param filePath
     * @param collectionArr
     * @param clazzArr
     * @param tag
     */
    public static void exportMutiToFile(String filePath, Collection<?>[] collectionArr, int[] tag) {
    	ExcelExportUtil.exportMutiSheetToFile(filePath,collectionArr,tag);
    } 
    
    /**
     * 读取excel文件为Map类型的列表
     * @param file
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
	
    public static List<Map<String, String>> importFromFile(File file) throws IOException, InvalidFormatException {
        return ExcelImportUtil.importFromFile(file);
    }

    /**
     * 
     * @param file
     * @param clazz
     * @return 读取excel文件为泛型类型的列表
     * @throws IllegalAccessException
     * @throws IntrospectionException
     * @throws InvalidFormatException
     * @throws IOException
     * @throws InstantiationException
     * @throws InvocationTargetException
     */
    public static <T> List<T> importFromFile(File file, Class<T> clazz) throws IllegalAccessException,
            IntrospectionException, InvalidFormatException, IOException, InstantiationException,
            InvocationTargetException {
        return ExcelImportUtil.importFromFile(file, clazz);
    }

}