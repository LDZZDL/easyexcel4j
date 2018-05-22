package example;

import com.github.easyexcel4j.reader.context.ReaderContext;
import com.github.easyexcel4j.reader.listener.ExcelReaderListener;

import java.util.List;

/**
 * @author LDZZDL
 * @create 2018-05-23 0:59
 **/
public class LargeSheetListener implements ExcelReaderListener {

    @Override
    public void invoke(List<String> datas, ReaderContext readerContext) {
        if(readerContext.getCurrentRowNumber() % 1000 == 0){
            System.out.println("当前行为：" + readerContext.getCurrentRowNumber() +
                readerContext.getCurrentSheetIndex() + "," +
                readerContext.isBlankRow());
        }
    }
}
