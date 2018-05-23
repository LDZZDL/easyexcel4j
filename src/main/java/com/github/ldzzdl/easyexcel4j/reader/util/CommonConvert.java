package com.github.ldzzdl.easyexcel4j.reader.util;

import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.beanutils.Converter;
import org.apache.poi.ss.usermodel.DateUtil;

import java.util.Date;

/**
 * @author LDZZDL
 * 读取Excel的通用转换器
 */
public class CommonConvert {

    /**
     * 向ConvertUtils注册类型转换器
     */
    public void register() {
        //String-Integer类型转换器
        ConvertUtils.register(new Converter() {
            @Override
            public <T> T convert(Class<T> aClass, Object o) {
                if (o != null && o instanceof String) {
                    int reuslt = (int) Double.parseDouble((String) o);
                    return (T) Integer.valueOf(reuslt);
                } else {
                    return null;
                }
            }
        }, Integer.class);
        //String-int类型转换器
        ConvertUtils.register(new Converter() {
            @Override
            public <T> T convert(Class<T> aClass, Object o) {
                if (o != null && o instanceof String) {
                    int reuslt = (int) Double.parseDouble((String) o);
                    return (T) Integer.valueOf(reuslt);
                } else {
                    return null;
                }
            }
        }, int.class);
        //String-Date类型转换器
        ConvertUtils.register(new Converter() {
            @Override
            public <T> T convert(Class<T> aClass, Object o) {
                if (o != null && o instanceof String) {
                    Date date = null;
                    try{
                        double result = Double.parseDouble((String) o);
                        date = DateUtil.getJavaDate(result);
                    }catch (Exception e){ }
                    return (T) date;
                } else {
                    return null;
                }
            }
        }, Date.class);
    }

}
