package com.github.easyexcel4j.reader.util;

import com.github.easyexcel4j.reader.context.ReaderContext;
import com.github.easyexcel4j.reader.listener.ExcelReaderListener;

import java.util.List;

/**
 * @author LDZZDL
 * @create 2018-05-22 17:13
 **/
public class ReadLargeSheet implements ExcelReaderListener {

    @Override
    public void invoke(List<String> datas, ReaderContext readerContext) {
        if(readerContext.getCurrentRowNumber() % 1000 == 0){
            System.out.println(readerContext.getCurrentRowNumber());
        }
    }
}
