# exzel

an easy java object to excel mapper framework based on poi with more features

# 一、功能介绍

基于apache poi框架，以注解形式定义java对象到excel对象的映射方式。

* [X] 支持多sheet页
* [X] 支持对象嵌套递归查找
* [X] 支持日期格式化
* [X] 支持自定义字段顺序
* [X] 支持自定义字段名，未设置则使用成员变量名
* [X] 支持单独设置每个表头字段的样式
* [X] 支持开启自动设置列宽(追踪列中最长值的长度)
* [X] 支持文本、数值、日期、布尔类型的数据，并以文本类型兜底
* [X] 支持Iterable接口的实现类映射时自动纵向拓展
* [X] 支持导出到输出流
* [X] 支持导出到httpServletResponse
* [x] 性能统计日志
* [x] 支持表头文本以方法形式提供
* [x] 支持域通过java方法计算结果作为单元格内容

# 二、实现

## 1.注解定义

```java
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
```

```java

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelRecursiveMapping {
    // 用于递归到对象内部扫描字段映射
}
```

```java
/**
 * excel表头单元格样式（仅表头行，支持按列自定义）
 *
 * @author gaattc
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
 
    IndexedColors backgroundColor() default IndexedColors.WHITE;
 
    FillPatternType fillPatternType() default FillPatternType.SOLID_FOREGROUND;
 
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.GENERAL;
 
    VerticalAlignment verticalAlignment() default VerticalAlignment.BOTTOM;
 
    IndexedColors fontColor() default IndexedColors.BLACK;
 
    short fontSize() default 12;
 
    boolean bold() default false;
 
    boolean italic() default false;
 
    FontUnderline underline() default FontUnderline.NONE;
 
}
```

## 2.excel对象生成器

ExcelGenerator

## 3.Excel输出器

ExcelExporter

# 三、使用示例

示例1：

```java

@SuppressWarnings("unused")
public class ExcelGeneratorTest extends TestCase {

    public void test() throws Exception {
        Foo source = new Foo();
        URI uri = getClass().getClassLoader().getResource("test.xlsx").toURI();
        Path path = Paths.get(uri);
        path.toFile().createNewFile();
        new ExcelExporter(source)
                .generate()
                .output(new BufferedOutputStream(Files.newOutputStream(path)));
    }

    private final static class Foo {
        @ExcelMapping(columnIndex = 0)
        private final String stringField = "stringField";
        @ExcelMapping(columnIndex = 1)
        @ExcelStyle(cellType = CellType.NUMERIC)
        private final int IntField = 233;
        @ExcelMapping(columnIndex = 3)
        @ExcelStyle(cellType = CellType.NUMERIC)
        private final long longField = 2333L;
        @ExcelMapping(columnIndex = 4)
        @ExcelStyle(cellType = CellType.BOOLEAN)
        private final boolean boolField = true;
        @ExcelMapping(columnIndex = 5, tryFormatDateTime = true)
        private final long dateField = System.currentTimeMillis();
        @ExcelMapping(sheetName = "iterable", columnIndex = 0)
        private final List<String> iterableField = CollectionUtils.list("1", "2", "3", "4");
        @ExcelRecursiveMapping
        private final Set<Bar> innerClassField = CollectionUtils.set(new Bar(), new Bar(), new Bar(), new Bar());
    }

    private final static class Bar {
        @ExcelMapping(columnIndex = 6)
        @ExcelStyle(autoSizeColumn = true)
        private final String innerStringField = "Bar#innerStringField";
        @ExcelRecursiveMapping
        private final Inner InnerClassField = new Inner();
    }

    private final static class Inner {
        @ExcelMapping(columnIndex = 2)
        @ExcelStyle(autoSizeColumn = true)
        private final String innerStringField = "Inner#innerStringField";
    }

}
```

示例2，增加对列名和单元格内容的计算：

```java
@Test
public void testConvert() throws Exception {
    TestForConvert source = new TestForConvert();
    Workbook workbook = new ExcelExporter(source)
            .generate()
            .getWorkbook();
    URI uri = getClass().getClassLoader().getResource("testConvert.xlsx").toURI();
    Path path = Paths.get(uri);
    path.toFile().createNewFile();
    new ExcelExporter(source)
            .generate()
            .output(Files.newOutputStream(path));
}
 
private final static class TestForConvert {
    // 优先使用Supplier
    @ExcelMapping(columnIndex = 0, contentConverter = "com.finebi.excel.ExcelGeneratorTest$TestForConvert#getValue")
    @ExcelStyle(columnName = "setColumnName", columnNameSupplier = "com.finebi.excel.ExcelGeneratorTest$TestForConvert#getColumnName")
    private final String javaFieldName = "originValue";
    // 其次使用设置的字段名
    @ExcelMapping(columnIndex = 1, contentConverter = "wrong express")
    @ExcelStyle(columnName = "setColumnName", columnNameSupplier = "wrong express")
    private final String javaFieldName1 = "originValue";
    // 最后默认使用javaFieldName
    @ExcelMapping(columnIndex = 2)
    @ExcelStyle(columnNameSupplier = "wrong express")
    private final String javaFieldName2 = "originValue";
 
    private static String getColumnName() {
        return "suppliedColumnName";
    }
 
    private static String getValue(String originValue) {
        return "convertedValue";
    }
 
}
```

# 四、性能报告

![](https://fastly.jsdelivr.net/gh/GAATTC0/MyPicGoOSS@main/img/flamegraph.png)

![](https://fastly.jsdelivr.net/gh/GAATTC0/MyPicGoOSS@main/img/Snipaste_2023-04-18_17-44-36.png)