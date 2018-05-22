package com.github.easyexcel4j.writer.util;

import com.github.easyexcel4j.metadata.ExcelType;

/**
 *Excel导出的上下文环境
 */
public class ExcelWriterContext {

    /**
     * font名称
     */
    private String fontName;
    /**
     * 字体大小
     */
    private short fontHeightInPoints;
    /**
     * excel的文件类型
     */
    private ExcelType excelType;

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public short getFontHeightInPoints() {
        return fontHeightInPoints;
    }

    public void setFontHeightInPoints(short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }


    public ExcelType getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelType excelType) {
        this.excelType = excelType;
    }
}
