package com.li.excel.row;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 
 * @Title: RowFactory.java 
 * @Package com.li.excel.row 
 * @Description: row工厂类 
 * @author leevan
 * @date 2018年11月14日 下午3:12:29
 * @version 1.0.0
 */
public final class RowFactory {

    private RowFactory() {
    }

    public static Row getOrCreate(Sheet sheet, int index) {
        if (sheet == null) {
            throw new NullPointerException("sheet is null !");
        }
        Row row = sheet.getRow(index);
        if (row == null) {
            return sheet.createRow(index);
        }
        return row;
    }
}
