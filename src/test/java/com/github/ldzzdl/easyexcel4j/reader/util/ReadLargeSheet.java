package com.github.ldzzdl.easyexcel4j.reader.util;

import com.github.ldzzdl.easyexcel4j.reader.context.ReaderContext;
import com.github.ldzzdl.easyexcel4j.reader.listener.ExcelReaderListener;

import java.util.List;

/**
 * @author LDZZDL
 * @create 2018-05-22 17:13
 **/
public class ReadLargeSheet implements ExcelReaderListener {

    @Override
    public void invoke(List<String> datas, ReaderContext readerContext) {
        if(readerContext.getCurrentRowNumber() % 10000 == 0){
            System.out.println("当前行为：" + readerContext.getCurrentRowNumber());
        }
    }
}
