package com.vike0906.be.help;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class XSSFExcelRead<T> implements ExcelRead<T> {

    @Override
    public List<T> excelToBean(String filePath, Class<T> clazz) throws IllegalAccessException, InstantiationException {

        FileInputStream fileInputStream = fileToFileInputStream(filePath);
        Workbook wb = null;
        try {
            wb = new XSSFWorkbook(fileInputStream);
            ReadHandle<T> readHandle = new ReadHandle<>();
            return readHandle.excelToBeans(wb, clazz);
        } catch (IOException e) {
            e.printStackTrace();
            throw new IllegalAccessException("IO异常");
        }finally {
            if(fileInputStream!=null){
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(wb != null){
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
