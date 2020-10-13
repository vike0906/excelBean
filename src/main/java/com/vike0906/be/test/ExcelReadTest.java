package com.vike0906.be.test;

import com.vike0906.be.help.ExcelRead;
import com.vike0906.be.help.ExcelReadHelp;

import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class ExcelReadTest {

    public static void main(String [] args){

        ExcelReadHelp<ExcelToBean> excelReadHelp = new ExcelReadHelp<>();

        try {

            ExcelRead<ExcelToBean> excelRead = excelReadHelp.createExcelBean(ExcelReadHelp.XSSF_TYPE);

            List<ExcelToBean> excelToBeans = excelRead.excelToBean("D:\\app\\excelToBeanTest.xlsx", ExcelToBean.class);

            for (ExcelToBean etb: excelToBeans) {
                System.out.println(etb.toString());
            }

        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        }

    }


}
