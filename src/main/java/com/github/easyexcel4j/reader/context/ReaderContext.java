package com.github.easyexcel4j.reader.context;

import com.github.easyexcel4j.metadata.ExcelType;

import java.util.ArrayList;
import java.util.List;

/**
 * @author LDZZDL
 * 导入Excel的上下文环境
 */
public class ReaderContext {

    /**
     * sheet的名称集合
     */
    private List<String> sheetNames;

    /**
     * 当前的sheet的序号（默认从0开始）
     */
    private int currentSheetIndex;

    /**
     * 当前行号（默认从0开始）
     */
    private int currentRowNumber;

    /**
     * 判断是否为空行
     */
    private boolean isBlankRow;

    /**
     * 当前行的最后一列的序号（默认从0开始）
     */
    private int lastColumnNumber;

    /**
     * Excel的文件类型
     */
    private ExcelType excelType;

    /**
     * Excel的模型数据
     */
    private Class clazz;

    public List<String> getSheetNames() {
        return sheetNames;
    }

    public void setSheetName(String sheetName) {
        this.sheetNames.add(sheetName);
    }

    public int getCurrentSheetIndex() {
        return currentSheetIndex;
    }

    public void setCurrentSheetIndex(int currentSheetIndex) {
        this.currentSheetIndex = currentSheetIndex;
    }

    public int getCurrentRowNumber() {
        return currentRowNumber;
    }

    public void setCurrentRowNumber(int currentRowNumber) {
        this.currentRowNumber = currentRowNumber;
    }

    public boolean isBlankRow() {
        return isBlankRow;
    }

    public void setBlankRow(boolean blankRow) {
        isBlankRow = blankRow;
    }

    public int getLastColumnNumber() {
        return lastColumnNumber;
    }

    public void setLastColumnNumber(int lastColumnNumber) {
        this.lastColumnNumber = lastColumnNumber;
    }

    public ExcelType getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelType excelType) {
        this.excelType = excelType;
    }

    public Class getClazz() {
        return clazz;
    }

    public void setClazz(Class clazz) {
        this.clazz = clazz;
    }

    @Override
    public String toString() {
        return "ReaderContext{" +
                "sheetNames=" + sheetNames +
                ", currentSheetIndex=" + currentSheetIndex +
                ", currentRowNumber=" + currentRowNumber +
                ", isBlankRow=" + isBlankRow +
                ", lastColumnNumber=" + lastColumnNumber +
                ", excelType=" + excelType +
                ", clazz=" + clazz +
                '}';
    }

    public ReaderContext() {
        this.sheetNames = new ArrayList<>();
        this.isBlankRow = false;
    }
}
