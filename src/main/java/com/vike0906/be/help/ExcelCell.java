package com.vike0906.be.help;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {


	/**
	 * 字段对应excel列序号
	 * @return
	 */
	int index() default 0;


	/**
	 * 字段对应标题
	 * @return
	 */
	String title() default "";


}
