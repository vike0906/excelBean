# excelBean
#### 一款基于poi4.1.2封装的excel与java bean转换工具
+ 主要实现方式是注解配合Java反射完成Beany与excel文件的赋值/取值
+ 主要解决excel文件导入与导出时需要手动给对应Bean赋值/取值，枯燥且代码量巨大
+ 使用方式
  - 读取excel文件时:
  ```
  ExcelReadHelp<ExcelToBean> excelReadHelp = new ExcelReadHelp<>();
        try {
            ExcelRead<ExcelToBean> excelRead = excelReadHelp.createExcelBean(ExcelReadHelp.XSSF_TYPE);
            List<ExcelToBean> excelToBeans = excelRead.excelToBean("导出文件.xlsx", ExcelToBean.class);
            for (ExcelToBean etb: excelToBeans) {
                System.out.println(etb.toString());
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        }
   ```
   - 导出excel文件时：
   ```
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
            excelWriteHelp.createBeanExcel("导出文件.xlsx",excelToBeans);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    ```
+ maven 依赖
```
    <dependencies>
        <!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>4.1.2</version>
        </dependency>
    </dependencies>
```


