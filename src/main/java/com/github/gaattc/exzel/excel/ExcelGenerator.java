package com.github.gaattc.exzel.excel;

import com.google.common.base.Stopwatch;
import com.google.common.base.Strings;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import com.sun.istack.internal.Nullable;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.helpers.MessageFormatter;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * poi excel对象生成器，格式为xlsx。
 * 支持将对象域映射到excel字段，配合{@link ExcelMapping}注解使用
 * 仅支持{@link CellType#STRING}、{@link CellType#NUMERIC}、{@link CellType#BOOLEAN}字段类型，默认为String
 *
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/13
 */
@Slf4j
public class ExcelGenerator {

    private static final int DEFAULT_FIELD_START_ROW = 0;
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
    private final Object source;

    public ExcelGenerator(Object source) {
        this.source = source;
    }

    public Workbook generate() throws Exception {
        Stopwatch stopwatch = Stopwatch.createStarted();
        dataBinding(source, DEFAULT_FIELD_START_ROW);
        transferToWorkbook();
        log.info("[ExcelGenerator] excel workbook generated successfully from {}, cost: {}",
                source.getClass().getSimpleName(),
                stopwatch.stop()
        );
        return workBook;
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
                startRow = bindInnerLevelField(source, startRow, field);
            }
        }
    }

    private void bindCurrentLevelField(Object source, Field field, ExcelMapping excelMapping, int startRow) throws IllegalAccessException {
        String sheetName = excelMapping.sheetName();
        Table<Integer, Integer, Object> sheet = WORKBOOK_DATA.computeIfAbsent(sheetName, i -> HashBasedTable.create());
        int columnIndex = excelMapping.columnIndex();
        String columnName = field.getName();
        boolean tryFormatDateTime = excelMapping.tryFormatDateTime();
        Object fieldData = field.get(source);
        if (isIterable(field.getType())) {
            Iterator<?> iterator = ((Iterable<?>) fieldData).iterator();
            for (int row = 0; iterator.hasNext(); row++) {
                Object data = tryFormatDateTime(iterator.next(), tryFormatDateTime);
                checkConflict(sheet, columnIndex, row, data);
                sheet.put(row, columnIndex, data);
            }
        } else {
            fieldData = tryFormatDateTime(fieldData, tryFormatDateTime);
            checkConflict(sheet, columnIndex, startRow, fieldData);
            sheet.put(startRow, columnIndex, fieldData);
        }
        // 表头字段样式
        ExcelStyle excelStyle = field.getAnnotation(ExcelStyle.class);
        if (null != excelStyle) {
            WORKBOOK_HEADER_STYLE.put(sheetName, columnIndex, excelStyle);
            String setColumnName = excelStyle.columnName();
            if (!Strings.isNullOrEmpty(setColumnName)) {
                columnName = setColumnName;
            }
        }
        // 字段名
        WORKBOOK_COLUMN_NAME.put(sheetName, columnIndex, columnName);
    }

    private int bindInnerLevelField(Object source, int startRow, Field field) throws IllegalAccessException {
        boolean innerIterable = isIterable(field.getType());
        if (innerIterable) {
            Iterable<?> iterableField = (Iterable<?>) field.get(source);
            for (Object fieldObj : iterableField) {
                dataBinding(fieldObj, startRow++);
            }
        } else {
            dataBinding(field.get(source), startRow);
        }
        return startRow;
    }

    private Object tryFormatDateTime(Object data, boolean tryFormatDateTime) {
        if (tryFormatDateTime && Long.class.isAssignableFrom(data.getClass())) {
            return DateFormatUtils.format((Long) data, PATTERN);
        } else {
            return data;
        }
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
                CellStyle cellStyle = createStyle(workBook, excelStyle);
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
                    setValueByType(cell, columnMapEntry.getValue(), columnStyleMap.get(columnNum));
                }
            }
            // 调整列宽
            for (Integer columnNum : WORKBOOK_COLUMN_NAME.row(sheetName).keySet()) {
                if (sheet.isColumnTrackedForAutoSizing(columnNum)) {
                    sheet.autoSizeColumn(columnNum);
                }
            }
        }
    }

    @Nullable
    private CellStyle createStyle(Workbook workbook, ExcelStyle excelStyle) {
        if (null == excelStyle) {
            return null;
        }
        CellStyle cellStyle = STYLE_CACHE.get(excelStyle);
        if (null != cellStyle) {
            return cellStyle;
        }
        CellStyle style = workbook.createCellStyle();
        // 设置填充色
        style.setFillForegroundColor(excelStyle.backgroundColor().index);
        style.setFillPattern(excelStyle.fillPatternType());
        // 设置对齐方式
        style.setAlignment(excelStyle.horizontalAlignment());
        style.setVerticalAlignment(excelStyle.verticalAlignment());
        // 字体样式
        Font font = workbook.createFont();
        // 字体颜色
        font.setColor(excelStyle.fontColor().index);
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

    private void setValueByType(SXSSFCell cell, Object value, ExcelStyle excelStyle) {
        if (null == excelStyle) {
            cell.setCellValue(value.toString());
            return;
        }
        // 目前仅支持数值、文本、布尔类型
        switch (excelStyle.cellType()) {
            case NUMERIC:
                cell.setCellValue(Long.parseLong(value.toString()));
                break;
            case BOOLEAN:
                cell.setCellValue(Boolean.parseBoolean(value.toString()));
                break;
            case STRING:
            default:
                cell.setCellValue(value.toString());
        }
    }

}
