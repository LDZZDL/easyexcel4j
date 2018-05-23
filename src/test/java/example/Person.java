package example;

import com.github.ldzzdl.easyexcel4j.annotation.Excel;

import java.util.Date;

/**
 * @author LDZZDL
 * @create 2018-05-22 21:14
 **/
public class Person{
    //2代表第2列
    @Excel(excelOrder = 2, excelTitle = "姓名")
    private String name;
    //3代表第三列
    @Excel(excelOrder = 3, excelTitle = "年龄")
    private Integer age;
    @Excel(excelOrder = 5, excelTitle = "爱好")
    private String hobby;
    @Excel(excelOrder = 6, excelTitle = "生日")
    private Date birthday;

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

    public String getHobby() {
        return hobby;
    }

    public void setHobby(String hobby) {
        this.hobby = hobby;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", hobby='" + hobby + '\'' +
                ", birthday=" + birthday +
                '}';
    }
}
