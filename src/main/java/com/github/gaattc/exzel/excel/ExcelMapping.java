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

    /**
     * 单元格内容转换器，需要提供静态方法的完整名，允许非public方法，如："com.finebi.excel.ExcelGeneratorTest$TestForConvert#getValue"
     * 且方法返回值需要与单元格类型匹配，否则转换无效，仍使用原数据
     * 与{@link ExcelMapping#tryFormatDateTime()}的生效先后顺序为先执行该方法再执行日期格式化。
     */
    String contentConverter() default "";

}
