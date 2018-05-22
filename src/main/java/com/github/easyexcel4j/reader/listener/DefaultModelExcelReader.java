package com.github.easyexcel4j.reader.listener;

import com.github.easyexcel4j.annotation.Excel;
import com.github.easyexcel4j.reader.context.ReaderContext;
import com.github.easyexcel4j.reader.util.CommonConvert;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author LDZZDL
 * 默认的模型监听器
 */
public class DefaultModelExcelReader implements ExcelReaderListener {

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
    private List datas;

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
        new CommonConvert().register();
    }

    public List getDatas() {
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
        if(checkSheetIndex(readerContext) && !readerContext.isBlankRow()){
            Class clazz = readerContext.getClazz();
            try {
                Object obj = clazz.newInstance();
                Field[] fields = clazz.getDeclaredFields();
                int size = datas.size();
                for (Field field : fields){
                    String fieldName = field.getName();
                    Excel annotation = field.getAnnotation(Excel.class);
                    int orderExcel = annotation.excelOrder() - 1;
                    if(size > orderExcel && datas.get(orderExcel) != null){
                        BeanUtils.setProperty(obj, fieldName, datas.get(annotation.excelOrder()-1));
                    }
                }
                this.datas.add(obj);
            } catch (IllegalAccessException | InstantiationException | InvocationTargetException e) {
                e.printStackTrace();
            }
        }
    }
}
