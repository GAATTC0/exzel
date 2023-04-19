package com.github.gaattc.exzel.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
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
     * 自动设置本列列宽
     */
    boolean autoSizeColumn() default false;

    IndexedColors backgroundColor() default IndexedColors.WHITE;

    FillPatternType fillPatternType() default FillPatternType.NO_FILL;

    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.GENERAL;

    VerticalAlignment verticalAlignment() default VerticalAlignment.BOTTOM;

    IndexedColors fontColor() default IndexedColors.BLACK;

    short fontSize() default 12;

    boolean bold() default false;

    boolean italic() default false;

    FontUnderline underline() default FontUnderline.NONE;

}
