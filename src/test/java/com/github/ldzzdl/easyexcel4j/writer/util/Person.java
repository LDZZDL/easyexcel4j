package com.github.ldzzdl.easyexcel4j.writer.util;

import com.github.ldzzdl.easyexcel4j.annotation.Excel;

/**
 * @author LDZZDL
 * @create 2018-05-22 16:45
 **/
public class Person {

    @Excel(excelOrder = 1, excelTitle = "姓名")
    private String name;
    @Excel(excelOrder = 2, excelTitle = "年龄")
    private Integer age;
    @Excel(excelOrder = 3, excelTitle = "运动")
    private String sport;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getSport() {
        return sport;
    }

    public void setSport(String sport) {
        this.sport = sport;
    }


    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", sport='" + sport + '\'' +
                '}';
    }
}
