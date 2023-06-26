package com.github.gaattc.exzel.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel表头单元格样式（仅表头行，支持按列自定义）
 *
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/13
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelStyle {

    CellType cellType() default CellType.STRING;

    String columnName() default "";

    /**
     * 表头字符串提供者，需要提供静态方法的完整名，允许非public方法，如："com.finebi.excel.ExcelGeneratorTest$TestForConvert#getColumnName"
     * 若与{@link ExcelStyle#columnName()}同时存在则优先使用该方法。
     */
    String columnNameSupplier() default "";

    /**
     * 自动设置本列列宽
     */
    boolean autoSizeColumn() default false;

    /**
     * 背景色，ARGB值
     */
    int backgroundColor() default 0xffffffff;

    FillPatternType fillPatternType() default FillPatternType.NO_FILL;

    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.GENERAL;

    VerticalAlignment verticalAlignment() default VerticalAlignment.BOTTOM;

    /**
     * 字体色，ARGB值
     */
    int fontColor() default 0xff000000;

    short fontSize() default 12;

    boolean bold() default false;

    boolean italic() default false;

    FontUnderline underline() default FontUnderline.NONE;

}
