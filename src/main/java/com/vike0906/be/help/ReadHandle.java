package com.vike0906.be.help;


import org.apache.poi.ss.usermodel.*;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 *
 */
public class ReadHandle<T> {


	/**
	 * 将excel数据转换为java bean集合
	 * @param wb
	 * @return
	 * @throws NullPointerException
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 */
	public List<T> excelToBeans(Workbook wb, Class<T> clazz) throws NullPointerException,InstantiationException,IllegalAccessException{

		if(wb==null){
			throw new NullPointerException("Workbook is null");
		}

		/*根据目标bean的注解获取字段与excel行数据对应关系 */
        List<FieldSort> fieldSorts = new ArrayList<>();
		Field[] declaredFields = clazz.getDeclaredFields();
        for(Field field:declaredFields){
            Annotation[] declaredAnnotations = field.getDeclaredAnnotations();
            for(Annotation annotation:declaredAnnotations){
                if(annotation instanceof ExcelCell){
                	ExcelCell excelCell = ((ExcelCell) annotation);
                    field.setAccessible(true);
                    FieldSort fieldSort = new FieldSort(field, excelCell.index());
                    fieldSorts.add(fieldSort);
                }
            }
        }

		fieldSorts.sort(Comparator.comparingInt(FieldSort::getIndex));

        List<T> list = new ArrayList<>();

        /*构造对象 */
		int numOfSheet = wb.getNumberOfSheets();
		for (int i = 0; i < numOfSheet; i++) {
			Sheet sheet = wb.getSheetAt(i);
			int lastRowNum = sheet.getLastRowNum();
			/*从第二行开始第一行是标题 */
			for (int j = 1; j <= lastRowNum; j++) {
				Row row = sheet.getRow(j);
				T t = cellToBean(row, fieldSorts, clazz.newInstance());
				list.add(t);
			}
		}
        return list;
    }

	/**
	 * 将excel行数据转为目标java write
	 * @param row
	 * @param fieldSorts
	 * @param t
	 * @return
	 * @throws IllegalAccessException
	 */
    private T cellToBean(Row row, List<FieldSort> fieldSorts, T t) throws IllegalAccessException {

		for (FieldSort fs:fieldSorts) {
			Field field = fs.getField();
			Cell cell = row.getCell(fs.getIndex());
			CellType cellType = cell.getCellType();
			if(cellType.equals(CellType._NONE)){
				throw new IllegalAccessException("cell is unknown type");
			}
			if(cellType.equals(CellType.ERROR)){
				throw new IllegalAccessException("cell is error");
			}
			switch (field.getGenericType().getTypeName()){
				case "java.lang.String":
					if(cellType.equals(CellType.STRING)){
						fs.getField().set(t, cell.getRichStringCellValue().getString());
					}else {
						fs.getField().set(t, null);
					}
					break;
				case "java.lang.Double":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, cell.getNumericCellValue());
					}else {
						fs.getField().set(t, null);
					}
					break;
				case "double":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, cell.getNumericCellValue());
					}else {
						fs.getField().set(t, 0D);
					}
					break;
				case "java.lang.Float":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, new Double(cell.getNumericCellValue()).floatValue());
					}else {
						fs.getField().set(t, null);
					}
					break;
				case "float":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, new Double(cell.getNumericCellValue()).floatValue());
					}else {
						fs.getField().set(t, 0F);
					}
					break;
				case "java.lang.Long":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, new Double(cell.getNumericCellValue()).longValue());
					}else {
						fs.getField().set(t, null);
					}
					break;
				case "long":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, new Double(cell.getNumericCellValue()).longValue());
					}else {
						fs.getField().set(t, 0L);
					}
					break;
				case "java.lang.Integer":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, new Double(cell.getNumericCellValue()).intValue());
					}else {
						fs.getField().set(t, null);
					}
					break;
				case "int":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, new Double(cell.getNumericCellValue()).intValue());
					}else {
						fs.getField().set(t, 0);
					}
					break;
				case "java.util.Date":
					if(cellType.equals(CellType.NUMERIC)){
						fs.getField().set(t, DateUtil.getJavaDate(cell.getNumericCellValue()));
					}else {
						fs.getField().set(t, null);
					}
					break;
				default:
					fs.getField().set(t, cell.getRichStringCellValue().getString());
					break;
			}


		}
		return t;
	}

}
