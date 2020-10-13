package com.vike0906.be.help;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public interface ExcelRead<T> {

    /**
     * 将excel文件解析成Java write
     * @param filePath
     * @param clazz
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    List<T> excelToBean(String filePath, Class<T> clazz) throws IllegalAccessException, InstantiationException;

    /**
     * 根据文件类型确定具体解析类
     * @param filePath
     * @return
     * @throws IllegalAccessException
     */
    default FileInputStream fileToFileInputStream(String filePath) throws IllegalAccessException{
        File file = new File(filePath);
        if(!file.exists()){
            throw new IllegalAccessException("文件不存在");
        }
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return fis;
    }
}
