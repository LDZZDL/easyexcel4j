package com.github.ldzzdl.easyexcel4j.reader.listener;

import com.github.ldzzdl.easyexcel4j.reader.context.ReaderContext;

import java.util.List;

/**
 * 读取Excel的监听器
 */
public interface ExcelReaderListener {

    /**
     * 返回每行的数据和导入的上下文环境
     * @param datas 每行的数据
     * @param readerContext 导入的上下文环境
     */
    void invoke(List<String> datas, ReaderContext readerContext);

}
