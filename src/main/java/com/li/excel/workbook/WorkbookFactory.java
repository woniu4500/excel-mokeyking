package com.li.excel.workbook;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @Title: WorkbookFactory.java 
 * @Package com.li.excel.workbook 
 * @Description: Workbook工厂 
 * @author leevan
 * @date 2018年11月14日 下午3:02:32
 * @version 1.0.0
 */
public final class WorkbookFactory {

    private WorkbookFactory() {
    }

    /**
     * 创建Workbook
     *
     * @return WorkbookContext
     */
    public static WorkbookContext createWorkbook() {
        Workbook workbook = new HSSFWorkbook();
        return new WorkbookContext(workbook);
    }

    public static WorkbookContext createWorkbook(WorkbookType workbookType) {

        //创建XLSX格式的excel
        if (workbookType == WorkbookType.XSSF) {
            Workbook workbook = new XSSFWorkbook();
            return new WorkbookContext(workbook);
        }

        if (workbookType == WorkbookType.SXSSF) {
            Workbook workbook = new SXSSFWorkbook();
            return new WorkbookContext(workbook);
        }
        //默认返回的Workbook
        return createWorkbook();
    }
    
    /**
     * 根据导入文件创建对应excel
     * @param inputStream
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    private static WorkbookContext createWorkbook(InputStream inputStream) throws IOException, InvalidFormatException {
        Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(inputStream);
        return new WorkbookContext(workbook);
    }
    
}