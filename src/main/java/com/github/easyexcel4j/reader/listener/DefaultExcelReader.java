package com.github.easyexcel4j.reader.listener;

import com.github.easyexcel4j.reader.context.ReaderContext;

import java.util.ArrayList;
import java.util.List;

/**
 * @author LDZZDL
 * 默认的读取监听器
 */
public class DefaultExcelReader implements ExcelReaderListener {

    /**
     * 开始读取的行数
     */
    private int startRowNumber;
    /**
     * 开始导入的sheet序号（从1开始）
     */
    private int startSheetNumber;
    /**
     * 结束导入的sheet序号
     */
    private int endSheetNumber;
    /**
     * 对应结果的集合
     */
    private List<String> datas;

    /**
     * 对传入的参数进行初始化
     * @param startRowNumber 开始读取的行数（默认从1开始）
     * @param startSheetNumber 始导入的sheet序号（从1开始）
     * @param endSheetNumber 结束导入的sheet序号（0代表，读取到末尾）
     */
    public void init(int startRowNumber, int startSheetNumber, int endSheetNumber){
        if(startRowNumber <= 0){
            this.startRowNumber = 1;
        }else{
            this.startRowNumber = startRowNumber;
        }
        if(startSheetNumber <= 0){
            this.startSheetNumber = 1;
        }else{
            this.startSheetNumber = startSheetNumber;
        }
        if(endSheetNumber < 0){
            this.endSheetNumber = 0;
        }else{
            this.endSheetNumber = endSheetNumber;
        }
        datas = new ArrayList<>();
    }

    public List<String> getDatas() {
        return datas;
    }

    /**
     * 检查当前的sheet序号
     * @param readerContext 导入Excel的上下文环境
     * @return 当前sheet是否合法
     */
    private boolean checkSheetIndex(ReaderContext readerContext){
        int currentSheetIndex = readerContext.getCurrentSheetIndex() + 1;
        if(currentSheetIndex >= startSheetNumber && endSheetNumber == 0 ||
                currentSheetIndex >= startSheetNumber && currentSheetIndex <= endSheetNumber) {
            int currentRowNumber = readerContext.getCurrentRowNumber() + 1;
            if (currentRowNumber >= startRowNumber) {
                return true;
            }
        }
        return false;
    }

    @Override
    public void invoke(List<String> datas, ReaderContext readerContext) {
       if (checkSheetIndex(readerContext)){
            int currentRowNumber = readerContext.getCurrentRowNumber() + 1;
            if(currentRowNumber >= startRowNumber){
                for(String data : datas){
                    this.datas.add(data);
                }
            }
        }
    }
}
