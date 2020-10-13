package com.vike0906.be.help;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class ExcelWriteHelp<T> {

    private static final String XSSF_SUFFIX = ".xlsx";

    public void createBeanExcel(String filePath, List<T> beans) throws IllegalAccessException{


        if(!filePath.endsWith(XSSF_SUFFIX)){
            throw new IllegalAccessException("当前只支持xlsx文件");
        }

        WriteHandle<T> writeHandle = new WriteHandle<>();

        Workbook workbook = writeHandle.beansToExcel(beans);

        FileOutputStream fileOut = null;

        try {
            File file = new File(filePath);
            if (!file.exists()) {
                file.createNewFile();
            }

            fileOut = new FileOutputStream(filePath);

            workbook.write(fileOut);

            fileOut.flush();

        } catch (Exception e) {
            throw new IllegalAccessException("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                if (null != workbook) {
                    workbook.close();
                }
            } catch (IOException e) {
                throw new IllegalAccessException("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }
    }
}
