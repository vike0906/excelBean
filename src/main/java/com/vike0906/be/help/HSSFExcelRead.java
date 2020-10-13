package com.vike0906.be.help;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class HSSFExcelRead<T> implements ExcelRead {

    @Override
    public List excelToBean(String filePath, Class clazz) throws IllegalAccessException, InstantiationException {
        FileInputStream fileInputStream = fileToFileInputStream(filePath);
        Workbook wb = null;
        try {
            wb = new HSSFWorkbook(fileInputStream);
        } catch (IOException e) {
            throw new IllegalAccessException("IO异常");
        }finally {
            if(wb != null){
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        ReadHandle<T> readHandle = new ReadHandle<>();
        return readHandle.excelToBeans(wb, clazz);
    }
}
