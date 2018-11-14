package com.li.excel.utils;

import java.beans.IntrospectionException;
import java.io.FileOutputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.li.excel.annotation.ExcelFieldInfo;
import com.li.excel.annotation.ExcelSheet;
import com.li.excel.row.RowContext;
import com.li.excel.sheet.SheetContext;
import com.li.excel.workbook.WorkbookContext;
import com.li.excel.workbook.WorkbookFactory;
import com.li.excel.workbook.WorkbookType;

/**
 * 
 * @Title: ExcelExportUtil.java 
 * @Package com.li.excel.utils 
 * @Description: 文件导出基础工具类 
 * @author leevan
 * @date 2018年11月14日 下午3:07:02
 * @version 1.0.0
 */
public class ExcelExportUtil {
    private static Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);
    
    private ExcelExportUtil() {
    	
    }

    /**
     * @param collection 集合
     * @param clazz      类
     * @param group        标记（0标记会忽略）
     * @return excel
     */
    public static byte[] writeSingleSheet(Collection<?> collection, int group, WorkbookType workbookType) throws
            NoSuchFieldException,
            IllegalAccessException, IntrospectionException, InvocationTargetException {
        //创建workbook上下文
        WorkbookContext workbookContext = WorkbookFactory.createWorkbook(workbookType);
        
        Class<?> clazz = collection.toArray()[0].getClass();
        List<ExcelFieldInfo> fieldInfoList = ReflectionUtil.getFieldInfoList(clazz);
        //处理标记，0则忽略
        handlegroup(group, fieldInfoList);

        //取得excel表头
        int m=fieldInfoList.size();
        List<String> headerList = new ArrayList<>();
        CellStyle[] fieldDataStyleArr = new CellStyle[m];
        CellStyle fieldDataStyle = workbookContext.getStyleConfiguration().getDefaultStyle(HorizontalAlignment.CENTER);
        for (int n=0;n<m;n++) {
            headerList.add(fieldInfoList.get(n).getName());
            HorizontalAlignment align = fieldInfoList.get(n).getAlign();
            //HorizontalAlignment align = HorizontalAlignment.LEFT;
            CellStyle cellStyle = workbookContext.toNativeWorkbook().createCellStyle();
            cellStyle.cloneStyleFrom(fieldDataStyle);
            cellStyle.setAlignment(align);
            fieldDataStyleArr[n]=cellStyle;
        }

        //创建表头
        ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        String sheetName = collection.getClass().getSimpleName();
        HSSFColorPredefined headColor = null;
        if (excelSheet != null) {
            if (excelSheet.name()!=null && excelSheet.name().trim().length()>0) {
                sheetName = excelSheet.name().trim();
            }
            headColor = excelSheet.headColor();
        }
        
        SheetContext sheetContext = workbookContext.createSheetAndHeader(0,sheetName,headColor,headerList 
        		.toArray(new String[headerList.size()]));
        //内容空，直接返回
        if (collection == null || collection.size() == 0) {
            return workbookContext.toNativeBytes();
        }

        //根据不同类型写row和cell
        for (Object object : collection) {
            RowContext rowContext = sheetContext.nextRow();
            for (int n=0;n<m;n++) {
                //值
                Object value = fieldInfoList.get(n).getMethod().invoke(object);
                int excelType = ReflectionUtil.getExcelTypeByObj(value);
                
                //处理需要特殊处理的字段，比如"第1周"
                String stringValue = handleSplicing(fieldInfoList.get(n), value, excelType);
                if (stringValue != null) {
                    value = stringValue;
                    //类型指定处理为字符串
                    excelType = ReflectionUtil.EXCEL_STRING_TYPE;
                }

                switch (excelType) {
                    case ReflectionUtil.EXCEL_NUMBER_TYPE:
                        rowContext.number((Number) value,fieldDataStyleArr[n]);
                        break;
                    case ReflectionUtil.EXCEL_DECIMAL_TYPE:
                        rowContext.decimal((Number) value, fieldInfoList.get(n).getFormat(),fieldDataStyleArr[n]);
                        break;
                    case ReflectionUtil.EXCEL_DATE_TYPE:
                        rowContext.date((Date) value, fieldInfoList.get(n).getFormat(),fieldDataStyleArr[n]);
                        break;
                    default:
                        //处理字符串和其他类型，对象为null则处理为默认值
                        String obj = String.valueOf((value == null
                                || StringUtils.isBlank(value.toString())) ? fieldInfoList.get(n).getDefaultValue() : value);
                        rowContext.text(obj,fieldDataStyleArr[n]);
                        break;
                }

                //宽度大于0 则设置宽度
                int width = fieldInfoList.get(n).getWidth();
                if (width > 0) {
                    rowContext.setColumnWidth(width);
                }

            }
        }
        return workbookContext.toNativeBytes();
    }
    
    
    /**
     * @param collection 集合
     * @param clazz      类
     * @param group        标记（0标记会忽略）
     * @return excel
     */
    public static byte[] writeMutiSheet(Collection<?>[] collectionArr, int[] groupArr, WorkbookType workbookType) throws
            NoSuchFieldException,
            IllegalAccessException, IntrospectionException, InvocationTargetException {

        //创建workbook上下文
        WorkbookContext workbookContext = WorkbookFactory.createWorkbook(workbookType);
        //内容空，直接返回
    	if(collectionArr == null || collectionArr.length ==0) {
            return workbookContext.toNativeBytes();
    	}
        //创建sheet
        for(int sheetNumber = 0; sheetNumber < collectionArr.length; sheetNumber++) {
            Class<?> clazz = collectionArr[sheetNumber].toArray()[0].getClass();
            List<ExcelFieldInfo> fieldInfoList = ReflectionUtil.getFieldInfoList(clazz);
            //处理标记，0则忽略
            handlegroup(groupArr[sheetNumber], fieldInfoList);

            //取得excel表头
            int m=fieldInfoList.size();
            List<String> headerList = new ArrayList<>();
            CellStyle[] fieldDataStyleArr = new CellStyle[m];
            CellStyle fieldDataStyle = workbookContext.getStyleConfiguration().getDefaultStyle(HorizontalAlignment.CENTER);
            for (int n=0;n<m;n++) {
                headerList.add(fieldInfoList.get(n).getName());
                HorizontalAlignment align = fieldInfoList.get(n).getAlign();
                //HorizontalAlignment align = HorizontalAlignment.LEFT;
                CellStyle cellStyle = workbookContext.toNativeWorkbook().createCellStyle();
                cellStyle.cloneStyleFrom(fieldDataStyle);
                cellStyle.setAlignment(align);
                fieldDataStyleArr[n]=cellStyle;
            }

            //创建表头
            ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
            String sheetName = collectionArr[sheetNumber].getClass().getSimpleName();
            HSSFColorPredefined headColor = null;
            if (excelSheet != null) {
                if (excelSheet.name()!=null && excelSheet.name().trim().length()>0) {
                    sheetName = excelSheet.name().trim();
                }
                headColor = excelSheet.headColor();
            }
            
            int i = 1;
            while(workbookContext.existSheet(sheetName)){
            	sheetName = sheetName.concat(String.valueOf(i));
            	i++;
            }
            
            SheetContext sheetContext = workbookContext.createSheetAndHeader(sheetNumber,sheetName,headColor,headerList 
                    .toArray(new String[headerList.size()]));

            //根据不同类型写row和cell
            for (Object object : collectionArr[sheetNumber]) {
                RowContext rowContext = sheetContext.nextRow();
                for (int n=0;n<m;n++) {
                    //值
                    Object value = fieldInfoList.get(n).getMethod().invoke(object);
                    int excelType = ReflectionUtil.getExcelTypeByObj(value);

                    //处理需要特殊处理的字段，比如"第3周"
                    String stringValue = handleSplicing(fieldInfoList.get(n), value, excelType);
                    if (stringValue != null) {
                        value = stringValue;
                        //类型指定处理为字符串
                        excelType = ReflectionUtil.EXCEL_STRING_TYPE;
                    }

                    switch (excelType) {
                        case ReflectionUtil.EXCEL_NUMBER_TYPE:
                            rowContext.number((Number) value,fieldDataStyleArr[n]);
                            break;
                        case ReflectionUtil.EXCEL_DECIMAL_TYPE:
                            rowContext.decimal((Number) value, fieldInfoList.get(n).getFormat(),fieldDataStyleArr[n]);
                            break;
                        case ReflectionUtil.EXCEL_DATE_TYPE:
                            rowContext.date((Date) value, fieldInfoList.get(n).getFormat(),fieldDataStyleArr[n]);
                            break;
                        default:
                            //处理字符串和其他类型，对象为null则处理为默认值
                            String obj = String.valueOf((value == null
                                    || StringUtils.isBlank(value.toString())) ? fieldInfoList.get(n).getDefaultValue() : value);
                            rowContext.text(obj,fieldDataStyleArr[n]);
                            break;
                    }

                    //宽度大于0 则设置宽度
                    int width = fieldInfoList.get(n).getWidth();
                    if (width > 0) {
                        rowContext.setColumnWidth(width);
                    }
                }
            }
        }
 
        return workbookContext.toNativeBytes();
    }

    /**
     * 处理group标记
     *
     * @param group           标记
     * @param fieldInfoList field信息列表
     */
    private static void handlegroup(int group, List<ExcelFieldInfo> fieldInfoList) {
        //group为0，表示忽略标记
        if (group != 0) {
            Iterator<ExcelFieldInfo> iterator = fieldInfoList.iterator();
            while (iterator.hasNext()) {
                ExcelFieldInfo fieldInfo = iterator.next();
                if (fieldInfo.getGroup().length == 1 && fieldInfo.getGroup()[0] == 0) {
                    continue;
                }
                //如果注解group里不包含参数group，则remove掉（不导出）
                boolean contain = false;
                for (int groupTemp : fieldInfo.getGroup()) {
                    if (groupTemp == group) {
                        contain = true;
                    }
                }
                //删除
                if (!contain) {
                    iterator.remove();
                }
            }
        }
    }

    //处理需要特殊处理的字段
    private static String handleSplicing(ExcelFieldInfo fieldInfo, Object value, int excelType) {
        String string = fieldInfo.getString();
        String format = fieldInfo.getFormat();
        //没有处理需求或者值为null则不处理
        if (StringUtils.isBlank(string) || value == null) {
            return null;
        }

        if (StringUtils.isBlank(format)) {
            return string.replace(ReflectionUtil.REPLACE_VALUE, value.toString());
        }
        if (excelType == ReflectionUtil.EXCEL_DATE_TYPE) {
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            return string.replace(ReflectionUtil.REPLACE_VALUE, sdf.format(value));
        }
        if (excelType == ReflectionUtil.EXCEL_DECIMAL_TYPE) {
            return NumberFormat.getInstance().format(value);
        }
        return null;
    }
    
    public static void exportSingleSheetToFile(String filePath,Collection<?> collection,int group) {
        FileOutputStream fileOutputStream = null;
        try {
        	WorkbookType workbookType = getWorkbookTypeFromPath(filePath);
			fileOutputStream = new FileOutputStream(filePath);
			fileOutputStream.write(writeSingleSheet(collection,group,workbookType));
            // flush
            fileOutputStream.flush();
		} catch (Exception e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        } finally {
            try {
                if (fileOutputStream!=null) {
                    fileOutputStream.close();
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
                throw new RuntimeException(e);
            }
        }
    }
    
    public static void exportMutiSheetToFile(String filePath,Collection<?>[] collectionArr,int[] groupArr) {
        FileOutputStream fileOutputStream = null;
    	WorkbookType workbookType = getWorkbookTypeFromPath(filePath);
        try {
			fileOutputStream = new FileOutputStream(filePath);
			fileOutputStream.write(writeMutiSheet(collectionArr,groupArr,workbookType));
            fileOutputStream.flush();
		} catch (Exception e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        } finally {
            try {
                if (fileOutputStream!=null) {
                    fileOutputStream.close();
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
                throw new RuntimeException(e);
            }
        }
    }

	public static WorkbookType getWorkbookTypeFromPath(String filePath) {
		WorkbookType workbookType = null;
		String [] filePathArr = filePath.split("\\.");
		if(!StringUtils.equals(filePathArr[filePathArr.length-1].toLowerCase(), "xls")
				||StringUtils.equals(filePathArr[filePathArr.length-1].toLowerCase(), "xlsx")) {
            throw new RuntimeException(">----------> 导出文件后缀名不正确,请检查是否为xls,xlsx格式");
		}
		if(StringUtils.equals(filePathArr[filePathArr.length-1].toLowerCase(), "xls")) {
			workbookType = WorkbookType.HSSF;
		}else {
			workbookType = WorkbookType.SXSSF;
		}
		return workbookType;
	}
	
}
