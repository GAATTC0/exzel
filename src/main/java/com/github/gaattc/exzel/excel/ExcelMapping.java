package com.github.gaattc.exzel.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/13
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelMapping {

    String sheetName() default "sheet";

    int columnIndex();

    /**
     * 若域为long类型，可以选择尝试格式化为可读日期 ExcelGenerator#PATTERN，格式化失败则使用string类型
     */
    boolean tryFormatDateTime() default false;

}
