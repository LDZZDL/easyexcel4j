package com.github.easyexcel4j.reader.listener;

import com.github.easyexcel4j.annotation.Excel;
import com.github.easyexcel4j.reader.context.ReaderContext;
import com.github.easyexcel4j.reader.util.CommonConvert;
import org.apache.commons.beanutils.BeanUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * @author LDZZDL
 * 根据标题名称自动封装成List数组的监听器
 */
public class AutoModelExcelReader implements ExcelReaderListener {

    /**
     * 是否开始读取数据
     */
    private boolean isDataRow;
    /**
     * 当前sheet的序号
     */
    private int currentSheetIndex;
    /**
     * 是否已经获取Class中相关信息
     */
    private boolean isInit;
    /**
     * 对应结果的集合
     */
    private List result;
    /**
     * 被注解属性的序号-对应Excel的列的序号
     */
    private Map<Integer, Integer> map;

    /**
     * Class类中被注解的属性值的名称
     */
    private List<String> fieldNames;
    /**
     * Class类中注解的excelTitle值
     */
    private List<String> titles;
    /**
     * 开始导入的sheet序号（从1开始）
     */
    private int startSheetNumber;
    /**
     * 结束导入的sheet序号
     */
    private int endSheetNumber;

    /**
     * 对传入的参数进行初始化
     * @param startSheetNumber 始导入的sheet序号（从1开始）
     * @param endSheetNumber 结束导入的sheet序号（0代表，读取到末尾）
     */
    public void init(int startSheetNumber, int endSheetNumber){
        //如果startSheetNumber的值小于0，默认从第一个sheet开始读取
        if(startSheetNumber <= 0){
            this.startSheetNumber = 1;
        }else{
            this.startSheetNumber = startSheetNumber;
        }
        //如果endSheetNumber的值小于0，默认读取到最后的sheet
        if(endSheetNumber < 0){
            this.endSheetNumber = 0;
        }else{
            this.endSheetNumber = endSheetNumber;
        }
    }

    @Override
    public void invoke(List<String> datas, ReaderContext readerContext) {
        if(!checkSheetIndex(readerContext)) return;
        boolean isBlankRow = readerContext.isBlankRow();
        if(isBlankRow) return;
        Class clazz = readerContext.getClazz();
        getAnnotationinformation(clazz);
        int currentSheetIndex = readerContext.getCurrentSheetIndex();
        if(this.currentSheetIndex != -1 && this.currentSheetIndex != currentSheetIndex){
            this.isDataRow = false;
        }
        if(!this.isDataRow){
            FindExcelOrderByTitle(datas, currentSheetIndex);
        }else{
            ModelConvert(datas, clazz);
        }

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
            return true;
        }
        return false;
    }

    /**
     * 对结果集合进行模型转换
     * @param datas 结果集合
     * @param clazz 模型的class
     */
    private void ModelConvert(List<String> datas, Class clazz) {
        try {
            Object obj = clazz.newInstance();
            int dataSize = datas.size();
            int titleSize = this.titles.size();
            for(int i = 0; i < titleSize; i++){
                Integer index = map.get(i);
                if(index != null){
                    if(dataSize > index && datas.get(index) != null){
                        try {
                            //封装成Javabean
                            BeanUtils.setProperty(obj, fieldNames.get(i), datas.get(index));
                        } catch (InvocationTargetException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
            this.result.add(obj);
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @param datas 结果集合
     * @param currentSheetIndex 当前sheet的序号
     */
    private void FindExcelOrderByTitle(List<String> datas, int currentSheetIndex) {
        int dataSize = datas.size();
        int titleSize = this.titles.size();
        for (int i = 0; i < titleSize; i++){
            for (int j = 0; j < dataSize; j++){
                String title = this.titles.get(i);
                String data = datas.get(j);
                if(data != null && title.equals(data.trim())){
                    map.put(i,j);
                    this.isDataRow = true;
                    this.currentSheetIndex = currentSheetIndex;
                }
            }
        }
    }

    /**
     * 获取Class类中被注解属性的相关信息
     * @param clazz Excel对应的模型
     */
    private void getAnnotationinformation(Class clazz) {
        if(!isInit && clazz != null){
            this.isInit = true;
            Field[] fields = clazz.getDeclaredFields();
            for(Field field : fields){
                Excel annotation = field.getAnnotation(Excel.class);
                if(annotation != null){
                    this.titles.add(annotation.excelTitle());
                    this.fieldNames.add(field.getName());
                }
            }
        }
    }

    public AutoModelExcelReader() {
        this.isDataRow = false;
        this.currentSheetIndex = -1;
        this.isInit = false;
        this.result = new ArrayList<>();
        this.map = new HashMap<>();
        this.fieldNames = new ArrayList<>();
        this.titles = new ArrayList<>();
        new CommonConvert().register();
    }

    public List getResult() {
        return result;
    }
}
