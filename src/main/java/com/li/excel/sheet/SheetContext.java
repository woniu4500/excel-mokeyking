package com.li.excel.sheet;

import org.apache.poi.ss.usermodel.Sheet;

import com.li.excel.row.RowContext;
import com.li.excel.row.RowFactory;
import com.li.excel.workbook.WorkbookContext;

/**
 * 
 * @Title: SheetContext.java 
 * @Package com.li.excel.sheet 
 * @Description: Sheet包装类
 * @author leevan
 * @date 2018年11月14日 下午3:08:27
 * @version 1.0.0
 */
public class SheetContext {

    private WorkbookContext workbookContext;
    private Sheet sheet;
    private RowContext currentRow;
    private int index = -1;

    public SheetContext(WorkbookContext workbookContext, Sheet sheet) {
        this.workbookContext = workbookContext;
        this.sheet = sheet;
    }

    public RowContext nextRow() {
        ++index;
        currentRow = null;
        return getCurrentRow();
    }

    public RowContext getCurrentRow() {
        if (index == -1) {
            return null;
        }
        if (currentRow == null) {
            currentRow = new RowContext(RowFactory.getOrCreate(sheet, index), this, workbookContext);
        }
        return currentRow;
    }

    public SheetContext setColumnWidths(int... columnNo) {
        for (int i = 0; i < columnNo.length; i++) {
            setColumnWidth(i, columnNo[i]);
        }
        return this;
    }

    public SheetContext hideGrid() {
        sheet.setDisplayGridlines(true);
        return this;
    }

    public SheetContext setColumnWidth(int columnNo, int width) {
        sheet.setColumnWidth(columnNo, width);
        return this;
    }

    public SheetContext skipOneRow() {
        return skipRows(1);
    }

    public SheetContext skipRows(int offset) {
        this.index += offset;
        return this;
    }
}