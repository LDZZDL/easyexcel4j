package com.github.ldzzdl.easyexcel4j.reader.listener;

import com.github.ldzzdl.easyexcel4j.reader.context.ReaderContext;

import java.util.ArrayList;
import java.util.List;

/**
 * 读取Excel的监听器的管理中心
 */
public class ExcelReaderListenerManager {

    protected List<ExcelReaderListener> excelReaderListeners = new ArrayList<>();

    public void registerListener(ExcelReaderListener listener){
        excelReaderListeners.add(listener);
    }

    public void removeListener(ExcelReaderListener listener){
        if(!excelReaderListeners.isEmpty()){
            excelReaderListeners.remove(listener);
        }
    }

    /**
     * 通知监听者，并返回每行的数据和导入的上下文环境
     * @param datas 每行的数据
     * @param readerContext 导入的上下文环境
     */
    public void notifyListener(List<String> datas, ReaderContext readerContext)  {
        for(ExcelReaderListener excelReaderListener : excelReaderListeners){
            excelReaderListener.invoke(datas, readerContext);
        }
    }
}
