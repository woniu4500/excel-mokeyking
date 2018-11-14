package com.li.excel.style;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 
 * @Title: StyleConfiguration.java 
 * @Package com.li.excel.style 
 * @Description: 样式配置类 
 * @author leevan
 * @date 2018年11月14日 下午3:07:45
 * @version 1.0.0
 */
public class StyleConfiguration {

    private Workbook workbook;//workbook


    private static final Byte DEFAULT_STYLE_KEY = 10;
    private static final Byte HEADER_STYLE_KEY = 11;
    private static final Byte DECIMAL_STYLE_KEY = 12;
    private static final Byte DATE_STYLE_KEY = 13;
    private static final Byte DATE_8_STYLE_KEY = 14;

    private static final Byte DATA_FORMAT_KEY = 20;

    private static final Byte FONT_KEY = 30;

    private final Map<Byte, CellStyle> buildInStyleMap = new HashMap<>(8);//内建样式
    private final Map<Byte, DataFormat> buildInFormatMap = new HashMap<>(2);//内建格式化
    private final Map<Byte, Font> buildInFontMap = new HashMap<>(2);//内建字体

    private final Map<String, CellStyle> customFormatStyleMap = new HashMap<>(8);//用户自定义格式化样式

    public StyleConfiguration(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * header样式
     *
     * @return CellStyle
     */
    public CellStyle getHeaderStyle(HSSFColor.HSSFColorPredefined colorPredefined) {

        if (buildInStyleMap.containsKey(HEADER_STYLE_KEY)) {
            return buildInStyleMap.get(HEADER_STYLE_KEY);
        }

        CellStyle headerStyle = workbook.createCellStyle();//头的样式
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置单元格的背景颜色为淡蓝色
        if(null!=colorPredefined) {
            headerStyle.setFillForegroundColor(colorPredefined.getIndex());
        }else {
            headerStyle.setFillForegroundColor(HSSFColorPredefined.PALE_BLUE.getIndex());
        }
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.setCommonStyle(headerStyle);
        buildInStyleMap.put(HEADER_STYLE_KEY, headerStyle);
        return headerStyle;
    }
        
    /**
     * 文本样式
     *
     * @return CellStyle
     */
    public CellStyle getTextStyle(CellStyle textStyle) {
        //text使用默认style
        return textStyle;
    }


    /**
     * 数字样式
     *
     * @return CellStyle
     */
    public CellStyle getNumberStyle(CellStyle numberStyle) {
        //number使用默认style
        return numberStyle;
    }


    /**
     * 小数格式
     *
     * @return CellStyle
     */
    public CellStyle getDecimalStyle(CellStyle decimalStyle) {
        if (buildInStyleMap.containsKey(DECIMAL_STYLE_KEY)) {
            return buildInStyleMap.get(DECIMAL_STYLE_KEY);
        }

        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        decimalStyle.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat("0.00"));
        this.setCommonStyle(decimalStyle);
        buildInStyleMap.put(DECIMAL_STYLE_KEY, decimalStyle);
        return decimalStyle;
    }

    /**
     * 日期样式 yyyy-MM-dd HH:mm
     *
     * @return CellStyle
     */
    public CellStyle getDateStyle(CellStyle dateStyle) {
        if (buildInStyleMap.containsKey(DATE_STYLE_KEY)) {
            return buildInStyleMap.get(DATE_STYLE_KEY);
        }

        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        dateStyle.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat("yyyy-MM-dd HH:mm"));
        this.setCommonStyle(dateStyle);
        buildInStyleMap.put(DATE_STYLE_KEY, dateStyle);
        return dateStyle;
    }

    /**
     * 日期样式 yyyy/MM/dd
     *
     * @return CellStyle
     */
    public CellStyle getDate8Style(CellStyle date8Style) {

        if (buildInStyleMap.containsKey(DATE_8_STYLE_KEY)) {
            return buildInStyleMap.get(DATE_8_STYLE_KEY);
        }

        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        date8Style.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat("yyyy/MM/dd"));
        this.setCommonStyle(date8Style);
        buildInStyleMap.put(DATE_8_STYLE_KEY, date8Style);
        return date8Style;
    }

    /**
     * 根据格式，创建返回样式对象
     *
     * @param format 格式
     * @return 样式对象
     */
    public CellStyle getCustomFormatStyle(String format,CellStyle customDateStyle) {

        //存在对应格式直接返回
        if (customFormatStyleMap.containsKey(format)) {
            return customFormatStyleMap.get(format);
        }
        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        customDateStyle.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat(format));
        this.setCommonStyle(customDateStyle);
        //放入map缓存
        customFormatStyleMap.put(format, customDateStyle);

        return customDateStyle;
    }

    /**
     * 默认样式,目前文字和整数使用的该样式
     *
     * @return CellStyle
     */
    public CellStyle getDefaultStyle(HorizontalAlignment align) {

        if (buildInStyleMap.containsKey(DEFAULT_STYLE_KEY)) {
            return buildInStyleMap.get(DEFAULT_STYLE_KEY);
        }

        CellStyle defaultStyle = workbook.createCellStyle();//默认样式
        defaultStyle.setAlignment(align);
        // 设置单元格边框为细线条
        this.setCommonStyle(defaultStyle);
        buildInStyleMap.put(DEFAULT_STYLE_KEY, defaultStyle);
        return defaultStyle;
    }

    /**
     * 设置通用的对齐居中、边框等
     *
     * @param style 样式
     */
    private void setCommonStyle(CellStyle style) {
        // 设置单元格居中对齐、自动换行
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        //设置单元格字体
        if (!buildInFontMap.containsKey(FONT_KEY)) {
            Font font = workbook.createFont();
            //通用字体
            font.setBold(true);
            font.setFontName("宋体");
            font.setFontHeight((short) 200);
            buildInFontMap.put(FONT_KEY, font);
        }
        style.setFont(buildInFontMap.get(FONT_KEY));

        // 设置单元格边框为细线条
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
    }
}