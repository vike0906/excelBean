package com.vike0906.be.help;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 *
 */
public class WriteHandle<T> {

	private static CellStyle DATE_FORMAT_CELL_STYLE = null;
	/**
	 * 将excel数据转换为java bean集合
	 * @param beans
	 * @return wb
	 * @throws NullPointerException
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 */
	public Workbook beansToExcel(List<T> beans) throws IllegalAccessException{

		if(beans==null || beans.size()==0){
			throw new IllegalAccessException("bean list is null or empty");
		}

		Workbook wb = new SXSSFWorkbook();

		Class<?> clazz = beans.get(0).getClass();

		/*根据目标bean的注解获取字段与excel行数据对应关系 */
        List<FieldSort> fieldSorts = new ArrayList<>();
		Field[] declaredFields = clazz.getDeclaredFields();
        for(Field field:declaredFields){
            Annotation[] declaredAnnotations = field.getDeclaredAnnotations();
            for(Annotation annotation:declaredAnnotations){
                if(annotation instanceof ExcelCell){
                	ExcelCell excelCell = ((ExcelCell) annotation);
                    field.setAccessible(true);
                    FieldSort fieldSort = new FieldSort(field, excelCell.index(), excelCell.title());
                    fieldSorts.add(fieldSort);
                }
            }
        }

		fieldSorts.sort(Comparator.comparingInt(FieldSort::getIndex));

		/* 处理表头 */
		String [] titles = new String[fieldSorts.size()];
		for(int i=0;i<fieldSorts.size();i++){
			titles[i] = fieldSorts.get(i).getTitle();
		}

		/* 生成Sheet表，并写入第一行*/
		Sheet sheet = buildDataSheet(wb, titles);

        /* 填充excel */
		int rowNum = 1;

		for (Iterator<T> it = beans.iterator(); it.hasNext();) {
			T data = it.next();
			if (data == null) {
				continue;
			}
			//输出行数据
			Row row = sheet.createRow(rowNum++);
			convertBeanToRow(fieldSorts, row, data);
		}

		if(DATE_FORMAT_CELL_STYLE!=null){
			DATE_FORMAT_CELL_STYLE = null;
		}

		return wb;
    }

	/**
	 * 将数据转换成行
	 * @param data 源数据
	 * @param row 行对象
	 * @return
	 */
	private void convertBeanToRow(List<FieldSort> fieldSorts, Row row, T data) throws IllegalAccessException {

		int cellNum = 0;

		Cell cell;

		for(FieldSort fs:fieldSorts){
			cell = row.createCell(cellNum++);
			Field field = fs.getField();
			switch (field.getGenericType().getTypeName()){
				case "java.lang.String":
					cell.setCellValue(null==field.get(data) ? "" : (String)fs.getField().get(data));
					break;
				case "java.lang.Double":
					cell.setCellValue(null== field.get(data) ? 0d : (Double) fs.getField().get(data));
					break;
				case "double":
					cell.setCellValue(null== field.get(data) ? 0d : (Double) fs.getField().get(data));
					break;
				case "java.lang.Float":
					cell.setCellValue(null== field.get(data) ? 0f : (Float) fs.getField().get(data));
					break;
				case "float":
					cell.setCellValue(null== field.get(data) ? 0f : (Float) fs.getField().get(data));
					break;
				case "java.lang.Long":
					cell.setCellValue(null== field.get(data) ? 0L : (Long) fs.getField().get(data));
					break;
				case "long":
					cell.setCellValue(null== field.get(data) ? 0L : (Long) fs.getField().get(data));
					break;
				case "java.lang.Integer":
					cell.setCellValue(null== field.get(data) ? 0 : (Integer) fs.getField().get(data));
					break;
				case "int":
					cell.setCellValue(null== field.get(data) ? 0 : (Integer) fs.getField().get(data));
					break;
				case "java.util.Date":

					if(DATE_FORMAT_CELL_STYLE==null){

						Workbook workbook = row.getSheet().getWorkbook();

						DATE_FORMAT_CELL_STYLE = workbook.createCellStyle();

						CreationHelper createHelper = workbook.getCreationHelper();

						short format = createHelper.createDataFormat().getFormat("yyyy-dd-MM");

						DATE_FORMAT_CELL_STYLE.setDataFormat(format);
					}

					cell.setCellValue(null== field.get(data) ? null : (Date) fs.getField().get(data));
					cell.setCellStyle(DATE_FORMAT_CELL_STYLE);
					break;
				default:
					cell.setCellValue("error");
					break;
			}
		}
	}

	/**
	 * 生成sheet表，并写入第一行数据（列头）
	 * @param workbook 工作簿对象
	 * @return 已经写入列头的Sheet
	 */
	private Sheet buildDataSheet(Workbook workbook, String [] titles) {
		Sheet sheet = workbook.createSheet();
		/* 设置列头宽度 */
		for (int i=0; i< titles.length; i++) {
			sheet.setColumnWidth(i, titles[i].length()*1024);
		}
		/* 设置默认行高 */
		sheet.setDefaultRowHeight((short) 400);
		/* 构建头单元格样式 */
		CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
		/* 写入第一行各列的数据 */
		Row head = sheet.createRow(0);
		for (int i = 0; i < titles.length; i++) {
			Cell cell = head.createCell(i);
			cell.setCellValue(titles[i]);
			cell.setCellStyle(cellStyle);
		}
		return sheet;
	}

	/**
	 * 设置第一行列头的样式
	 * @param workbook 工作簿对象
	 * @return 单元格样式对象
	 */
	private CellStyle buildHeadCellStyle(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		/* 对齐方式设置 */
		style.setAlignment(HorizontalAlignment.CENTER);
		/* 边框颜色和宽度设置 */
		style.setBorderBottom(BorderStyle.THIN);
		/* 下边框 */
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		/* 左边框 */
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(BorderStyle.THIN);
		/* 右边框 */
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		/* 上边框 */
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		/* 设置背景颜色 */
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		/* 粗体字设置 */
		Font font = workbook.createFont();
		font.setBold(true);
		style.setFont(font);
		return style;
	}

}
