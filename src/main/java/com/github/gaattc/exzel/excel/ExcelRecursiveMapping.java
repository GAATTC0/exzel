package com.github.gaattc.exzel.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标记域本身不映射到excel字段，而是向其内部继续寻找被{@link ExcelMapping}标记的域
 * 若域本身实现了Iterable接口，则映射到excel中纵向扩展
 *
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/13
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelRecursiveMapping {

}
