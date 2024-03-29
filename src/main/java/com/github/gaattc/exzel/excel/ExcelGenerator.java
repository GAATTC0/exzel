package com.github.gaattc.exzel.excel;

import com.google.common.base.Stopwatch;
import com.google.common.base.Strings;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import com.sun.istack.internal.NotNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.helpers.MessageFormatter;

import java.awt.Color;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * poi excel对象生成器，格式为xlsx。
 * 支持将对象域映射到excel字段，配合{@link ExcelRecursiveMapping}、{@link ExcelMapping}、{@link ExcelStyle}注解使用
 * 仅支持{@link CellType#STRING}、{@link CellType#NUMERIC}、{@link CellType#BOOLEAN}字段类型，默认为String
 *
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/13
 */
@Slf4j
public class ExcelGenerator {

    private static final int DEFAULT_FIELD_START_ROW = 0;
    private static final String NULL = "";
    private static final String ZERO_TIME_REPLACE = "--";
    private static final String PATTERN = "yyyy-MM-dd HH:mm:ss";
    /**
     * Map<sheetName, Table<row, column, data>>
     */
    private final Map<String, Table<Integer, Integer, Object>> WORKBOOK_DATA = new HashMap<>();
    /**
     * Table<sheetName, cloNum, ExcelStyle>
     */
    private final Table<String, Integer, ExcelStyle> WORKBOOK_HEADER_STYLE = HashBasedTable.create();
    /**
     * Table<sheetName, cloNum, cloName>
     */
    private final Table<String, Integer, String> WORKBOOK_COLUMN_NAME = HashBasedTable.create();
    private final Map<ExcelStyle, CellStyle> STYLE_CACHE = new HashMap<>();
    private final SXSSFWorkbook workBook = new SXSSFWorkbook();
    private final XSSFCellStyle dataRowStyleOdd = ((XSSFCellStyle) workBook.createCellStyle());
    private final XSSFCellStyle dataRowStyleEven = ((XSSFCellStyle) workBook.createCellStyle());
    private final Object source;
    private final ClassLoader classLoader;

    public ExcelGenerator(Object source) {
        this.source = source;
        classLoader = this.getClass().getClassLoader();
    }

    public ExcelGenerator(Object source, ClassLoader classLoader) {
        this.source = source;
        this.classLoader = classLoader;
    }

    public Workbook generate() throws Exception {
        Stopwatch stopwatch = Stopwatch.createStarted();
        createDataRowStyle();
        dataBinding(source, DEFAULT_FIELD_START_ROW);
        transferToWorkbook();
        log.info("[ExcelGenerator] excel workbook generated successfully from {}, cost: {}",
                source.getClass().getSimpleName(),
                stopwatch.stop()
        );
        return workBook;
    }

    private void createDataRowStyle() {
        dataRowStyleOdd.setFillForegroundColor(new XSSFColor(new Color(204, 232, 255)));
        dataRowStyleOdd.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        dataRowStyleEven.setFillForegroundColor(new XSSFColor(new Color(239, 243, 252)));
        dataRowStyleEven.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void dataBinding(Object source, int startRow) throws IllegalAccessException {
        Field[] fields = source.getClass().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            // 当前层级域映射
            ExcelMapping excelMapping = field.getAnnotation(ExcelMapping.class);
            if (null != excelMapping) {
                bindCurrentLevelField(source, field, excelMapping, startRow);
                continue;
            }
            // 下一层级域映射
            ExcelRecursiveMapping recursiveMapping = field.getAnnotation(ExcelRecursiveMapping.class);
            if (null != recursiveMapping) {
                bindInnerLevelField(source, startRow, field);
            }
        }
    }

    private void bindCurrentLevelField(Object source, Field field, ExcelMapping excelMapping, int startRow) throws IllegalAccessException {
        String sheetName = excelMapping.sheetName();
        Table<Integer, Integer, Object> sheet = WORKBOOK_DATA.computeIfAbsent(sheetName, i -> HashBasedTable.create());
        int columnIndex = excelMapping.columnIndex();
        Object fieldData = field.get(source);
        if (isIterable(field.getType()) && fieldData != null) {
            Iterator<?> iterator = ((Iterable<?>) fieldData).iterator();
            for (int row = 0; iterator.hasNext(); row++) {
                Object data = calculateData(iterator.next(), excelMapping);
                checkConflict(sheet, columnIndex, row, data);
                sheet.put(row, columnIndex, data);
            }
        } else {
            fieldData = calculateData(fieldData, excelMapping);
            checkConflict(sheet, columnIndex, startRow, fieldData);
            sheet.put(startRow, columnIndex, fieldData);
        }
        // 表头字段样式
        ExcelStyle excelStyle = field.getAnnotation(ExcelStyle.class);
        if (null != excelStyle) {
            WORKBOOK_HEADER_STYLE.put(sheetName, columnIndex, excelStyle);
        }
        // 字段名
        String columnName = generateColumnName(field.getName(), excelStyle);
        WORKBOOK_COLUMN_NAME.put(sheetName, columnIndex, columnName);
    }

    private void bindInnerLevelField(Object source, int startRow, Field field) throws IllegalAccessException {
        boolean innerIterable = isIterable(field.getType());
        if (innerIterable) {
            Iterable<?> iterableField = (Iterable<?>) field.get(source);
            int innerRow = startRow;
            for (Object fieldObj : iterableField) {
                dataBinding(fieldObj, innerRow++);
            }
        } else {
            dataBinding(field.get(source), startRow);
        }
    }

    /**
     * 计算出转换后的实际值
     */
    private Object calculateData(Object data, @NotNull ExcelMapping excelMapping) {
        data = nullless(data);
        // 先进行转换计算
        String converter = excelMapping.contentConverter();
        if (!Strings.isNullOrEmpty(converter)) {
            data = ReflectCaller.function(converter, data, classLoader);
        }
        // 最后尝试格式化日期
        if (excelMapping.tryFormatDateTime() && Long.class.isAssignableFrom(data.getClass())) {
            long longData = (long) data;
            if (longData != 0L) {
                return DateFormatUtils.format(longData, PATTERN);
            }
            // 为0则不展示为197001010800，而是占位符
            return ZERO_TIME_REPLACE;
        } else {
            return data;
        }
    }

    private Object nullless(Object data) {
        return null == data ? NULL : data;
    }

    private void checkConflict(Table<Integer, Integer, Object> sheet, int columnIndex, int row, Object data) {
        Object conflictValue = sheet.get(row, columnIndex);
        if (null != conflictValue) {
            throw new IllegalArgumentException(
                    MessageFormatter.arrayFormat("value conflict, check annotation if correct, row: {}, column: {}, value1: {}, value2: {}",
                            new Object[]{row, columnIndex, conflictValue, data}).getMessage()
            );
        }
    }

    private String generateColumnName(String fieldName, ExcelStyle excelStyle) {
        if (null != excelStyle) {
            String columnNameSupplier = excelStyle.columnNameSupplier();
            String setColumnedName = excelStyle.columnName();
            // 优先使用supplier提供的字段名
            if (!Strings.isNullOrEmpty(columnNameSupplier)) {
                String columnName = ReflectCaller.supplier(columnNameSupplier, classLoader);
                if (!Strings.isNullOrEmpty(columnName)) {
                    return columnName;
                }
            }
            // supplier提供了空字段名或计算错误，则降级使用设置的字段名
            if (!Strings.isNullOrEmpty(setColumnedName)) {
                return setColumnedName;
            }
        }
        return fieldName;
    }

    private void transferToWorkbook() {
        for (Map.Entry<String, Table<Integer, Integer, Object>> sheetMapEntry : WORKBOOK_DATA.entrySet()) {
            String sheetName = sheetMapEntry.getKey();
            SXSSFSheet sheet = workBook.createSheet(sheetName);
            // 表头
            SXSSFRow headerRow = sheet.createRow(DEFAULT_FIELD_START_ROW);
            Map<Integer, ExcelStyle> columnStyleMap = WORKBOOK_HEADER_STYLE.row(sheetName);
            for (Map.Entry<Integer, String> headerColumnNameMapEntry : WORKBOOK_COLUMN_NAME.row(sheetName).entrySet()) {
                Integer columnNum = headerColumnNameMapEntry.getKey();
                ExcelStyle excelStyle = columnStyleMap.get(columnNum);
                SXSSFCell headerRowCell = headerRow.createCell(columnNum);
                CellStyle cellStyle = createStyle(excelStyle);
                if (null != cellStyle) {
                    headerRowCell.setCellStyle(cellStyle);
                    // 设置自动列宽追踪
                    if (excelStyle.autoSizeColumn()) {
                        sheet.trackColumnForAutoSizing(columnNum);
                    }
                }
                headerRowCell.setCellValue(headerColumnNameMapEntry.getValue());
            }
            // 数据
            Map<Integer, Map<Integer, Object>> rowMap = sheetMapEntry.getValue().rowMap();
            for (Map.Entry<Integer, Map<Integer, Object>> rowMapEntry : rowMap.entrySet()) {
                // 在表头行下面开始写数据
                SXSSFRow row = sheet.createRow(rowMapEntry.getKey() + 1);
                // 行遍历
                for (Map.Entry<Integer, Object> columnMapEntry : rowMapEntry.getValue().entrySet()) {
                    Integer columnNum = columnMapEntry.getKey();
                    SXSSFCell cell = row.createCell(columnNum);
                    setCellStyle(cell, rowMapEntry.getKey());
                    setValueByType(cell, columnMapEntry.getValue(), columnStyleMap.get(columnNum));
                }
            }
            // 调整列宽
            adjustColumnSize(sheetName, sheet);
        }
    }

    private CellStyle createStyle(ExcelStyle excelStyle) {
        if (null == excelStyle) {
            return null;
        }
        CellStyle cellStyle = STYLE_CACHE.get(excelStyle);
        if (null != cellStyle) {
            return cellStyle;
        }
        XSSFCellStyle style = ((XSSFCellStyle) workBook.createCellStyle());
        // 设置填充色
        style.setFillForegroundColor(new XSSFColor(new Color(excelStyle.backgroundColor(), true)));
        style.setFillPattern(excelStyle.fillPatternType());
        // 设置对齐方式
        style.setAlignment(excelStyle.horizontalAlignment());
        style.setVerticalAlignment(excelStyle.verticalAlignment());
        // 字体样式
        XSSFFont font = ((XSSFFont) workBook.createFont());
        // 字体颜色
        font.setColor(new XSSFColor(new Color(excelStyle.fontColor(), true)));
        // 字体大小
        font.setFontHeightInPoints(excelStyle.fontSize());
        // 粗体
        font.setBold(excelStyle.bold());
        // 斜体
        font.setItalic(excelStyle.italic());
        // 下划线
        font.setUnderline(excelStyle.underline().getByteValue());
        style.setFont(font);
        STYLE_CACHE.put(excelStyle, style);
        return style;
    }

    private static boolean isIterable(Class<?> clazz) {
        return Iterable.class.isAssignableFrom(clazz);
    }

    private void setCellStyle(SXSSFCell cell, Integer rowNum) {
        if (rowNum % 2 == 0) {
            cell.setCellStyle(dataRowStyleEven);
        } else {
            cell.setCellStyle(dataRowStyleOdd);
        }
    }

    private void setValueByType(SXSSFCell cell, Object value, ExcelStyle excelStyle) {
        if (null == excelStyle) {
            cell.setCellValue(value.toString());
            return;
        }
        // 目前仅支持数值、文本、布尔类型
        switch (excelStyle.cellType()) {
            case NUMERIC:
                String decimal = new BigDecimal(value.toString())
                        // todo 这里的特性考虑要不要开放给外部自定
                        .setScale(2, RoundingMode.HALF_UP)
                        .stripTrailingZeros()
                        .toPlainString();
                cell.setCellValue(decimal);
                break;
            case BOOLEAN:
                cell.setCellValue(Boolean.parseBoolean(value.toString()));
                break;
            case STRING:
            default:
                cell.setCellValue(value.toString());
        }
    }

    private void adjustColumnSize(String sheetName, SXSSFSheet sheet) {
        for (Integer columnNum : WORKBOOK_COLUMN_NAME.row(sheetName).keySet()) {
            if (sheet.isColumnTrackedForAutoSizing(columnNum)) {
                sheet.autoSizeColumn(columnNum);
            }
        }
    }

}
