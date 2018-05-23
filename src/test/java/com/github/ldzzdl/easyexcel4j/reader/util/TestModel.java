package com.github.ldzzdl.easyexcel4j.reader.util;

import com.github.ldzzdl.easyexcel4j.annotation.Excel;

import java.util.Date;

/**
 * @author LDZZDL
 * @create 2018-05-22 16:45
 **/
public class TestModel {

    @Excel(excelOrder = 1, excelTitle = "姓名")
    private String name;
    @Excel(excelOrder = 2, excelTitle = "年龄")
    private int age;
    @Excel(excelOrder = 3, excelTitle = "年龄1")
    private Integer age1;
    @Excel(excelOrder = 4, excelTitle = "运动")
    private String sport;
    @Excel(excelOrder = 5, excelTitle = "生日")
    private Date birthday;
    @Excel(excelOrder = 6, excelTitle = "金钱")
    private double money;
    @Excel(excelOrder = 7, excelTitle = "剩余金钱")
    private Double rest;
    @Excel(excelOrder = 8, excelTitle = "结婚")
    private boolean isMaried;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Integer getAge1() {
        return age1;
    }

    public void setAge1(Integer age1) {
        this.age1 = age1;
    }

    public String getSport() {
        return sport;
    }

    public void setSport(String sport) {
        this.sport = sport;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public double getMoney() {
        return money;
    }

    public void setMoney(double money) {
        this.money = money;
    }

    public Double getRest() {
        return rest;
    }

    public void setRest(Double rest) {
        this.rest = rest;
    }

    public boolean isMaried() {
        return isMaried;
    }

    public void setMaried(boolean maried) {
        isMaried = maried;
    }

    @Override
    public String toString() {
        return "TestModel{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", age1=" + age1 +
                ", sport='" + sport + '\'' +
                ", birthday=" + birthday +
                ", money=" + money +
                ", rest=" + rest +
                ", isMaried=" + isMaried +
                '}';
    }
}
