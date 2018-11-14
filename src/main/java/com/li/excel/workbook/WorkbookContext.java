package com.li.excel.workbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.li.excel.row.RowContext;
import com.li.excel.sheet.SheetContext;
import com.li.excel.style.StyleConfiguration;

/**
 * 
 * @Title: WorkbookContext.java 
 * @Package com.li.excel.workbook 
 * @Description: Workbook包装类 
 * @author leevan
 * @date 2018年11月14日 下午3:03:25
 * @version 1.0.0
 */
public class WorkbookContext {

    private final StyleConfiguration styleConfiguration;
    private final Workbook workbook;

    WorkbookContext(Workbook workbook) {
        this.workbook = workbook;
        this.styleConfiguration = new StyleConfiguration(workbook);
    }

    /**
     * 创建Sheet
     *
     * @param sheetName sheet名
     */
    public SheetContext createSheet(String sheetName) {
        return new SheetContext(this, workbook.createSheet(sheetName));
    }
    
    public boolean existSheet(String sheetName) {
    	return workbook.getSheet(sheetName)==null?false:true;
    }

    /**
     * 创建名称为空的Sheet
     */
    public SheetContext createSheet() {
        return new SheetContext(this, workbook.createSheet());
    }
    
    public SheetContext createSheet(int sheetNumber,String sheetName){
    	Sheet sheet = workbook.createSheet();
    	workbook.setSheetName(sheetNumber, sheetName);
    	return new SheetContext(this, sheet);
    }

    
    /**
     * 创建名称为空的Sheet并生成header行
     *
     * @param headerArray 表头数组
     */
    public SheetContext createSheetAndHeader(HSSFColor.HSSFColorPredefined colorPredefined,String... headerArray) {
    	SheetContext sheetContext = createSheet();
        RowContext headerContext = sheetContext.nextRow();
        for (int i = 0; i < headerArray.length; i++) {
            headerContext.header(headerArray[i],colorPredefined);
        }

        return sheetContext;
    }
    
    public SheetContext createSheetAndHeader(int sheetNumber,String sheetName,HSSFColor.HSSFColorPredefined colorPredefined,String... headerArray) {
    	Sheet sheet = workbook.createSheet();
    	workbook.setSheetName(sheetNumber, sheetName);
        SheetContext sheetContext = new SheetContext(this, sheet);
        RowContext headerContext = sheetContext.nextRow();
        for (int i = 0; i < headerArray.length; i++) {
            headerContext.header(headerArray[i],colorPredefined);
        }
        return sheetContext;
    }

    /**
     * 原生Bytes
     */
    public byte[] toNativeBytes() {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException("ToNativeBytes Failed ", e);
        }
    }

    public StyleConfiguration getStyleConfiguration() {
        return this.styleConfiguration;
    }

    /**
     * 原生WorkBook
     */
    public Workbook toNativeWorkbook() {
        return workbook;
    }
}