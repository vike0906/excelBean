package com.vike0906.be.test;

import com.vike0906.be.help.ExcelWriteHelp;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/13
 */
public class ExcelWriteTest {

    public static void main(String [] args){

        ExcelWriteHelp<ExcelToBean> excelWriteHelp = new ExcelWriteHelp<>();

        List<ExcelToBean> excelToBeans = new ArrayList<>();

        for (int i=0; i<10; i++) {
            ExcelToBean etb = new ExcelToBean();
            etb.setName("姓名"+i);
            etb.setAge(20+i);
            etb.setSalary(1234.6f+i);
            etb.setBirthday(new Date(System.currentTimeMillis()));
            excelToBeans.add(etb);
        }

        try {
            excelWriteHelp.createBeanExcel("D:\\app\\beanToExcelTest.xlsx",excelToBeans);

        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

    }
}
