package com.vike0906.be.help;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class ExcelReadHelp<T> {

    public static final WorkbookType XSSF_TYPE = new WorkbookType("xlsx");
    public static final WorkbookType HSSF_TYPE = new WorkbookType("xls");

    public ExcelRead<T> createExcelBean(WorkbookType type) throws IllegalAccessException{

        if(XSSF_TYPE.equals(type)){

            return new XSSFExcelRead<>();

        }else if(HSSF_TYPE.equals(type)){

            return new HSSFExcelRead<>();

        }else {

            throw new IllegalAccessException("非法调用");
        }
    }

    static class WorkbookType{

        private String type;

        private WorkbookType(String type){
            this.type = type;
        }

        @Override
        public boolean equals(Object anObject){
            if (this == anObject) {
                return true;
            }
            if (anObject instanceof WorkbookType) {
                WorkbookType anotherWorkbookType = (WorkbookType)anObject;
                if(this.type.equals(anotherWorkbookType.type)){
                    return true;
                }
            }
            return false;
        }
    }
}
