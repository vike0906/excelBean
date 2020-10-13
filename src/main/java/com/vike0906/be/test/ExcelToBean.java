package com.vike0906.be.test;

import com.vike0906.be.help.ExcelCell;

import java.util.Date;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class ExcelToBean {

    @ExcelCell(title = "名称")
    private String name;

    @ExcelCell(index = 1, title = "年龄")
    private int age;

    @ExcelCell(index = 2, title = "薪资")
    private float salary;

    @ExcelCell(index = 3, title = "出生日期")
    private Date birthday;

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

    public float getSalary() {
        return salary;
    }

    public void setSalary(float salary) {
        this.salary = salary;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    @Override
    public String toString() {
        return "ExcelToBean{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", salary=" + salary +
                ", birthday=" + birthday +
                '}';
    }
}
