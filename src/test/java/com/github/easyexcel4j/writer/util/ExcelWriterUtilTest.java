package com.github.easyexcel4j.writer.util;

import com.github.easyexcel4j.metadata.ExcelType;
import org.junit.Test;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriterUtilTest {

    @Test
    public void test1() throws InvocationTargetException, NoSuchMethodException, IllegalAccessException, IOException {
        ExcelWriterUtil excelWriterUtil = new ExcelWriterUtil();
        List<Person> people = new ArrayList<>();
        Person person = new Person();
        person.setAge(12);
        person.setName("jack");
        person.setSport("篮球");
        people.add(person);
        people.add(person);
        people.add(person);
        ExcelWriterContext excelWriterContext = new ExcelWriterContext();
        excelWriterContext.setFontName("宋体");
        excelWriterContext.setFontHeightInPoints((short)11);
        excelWriterContext.setExcelType(ExcelType.XLSX);
        excelWriterUtil.writeModelList2Excel(people, excelWriterContext);
    }

}