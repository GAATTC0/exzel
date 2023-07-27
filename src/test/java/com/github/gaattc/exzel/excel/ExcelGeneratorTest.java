package com.github.gaattc.exzel.excel;

import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Ignore;
import org.junit.Test;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

/**
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/17
 */
@SuppressWarnings("unused")
public class ExcelGeneratorTest {

    @Test
    public void test() throws Exception {
        Foo source = new Foo();
        URI uri = getClass().getClassLoader().getResource("test.xlsx").toURI();
        Path path = Paths.get(uri);
        path.toFile().createNewFile();
        Workbook workbook = new ExcelExporter(source)
                .generate()
                .getWorkbook();
        URI expectUri = getClass().getClassLoader().getResource("expect.xlsx").toURI();
        BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(Paths.get(expectUri)));
        Workbook expectWorkbook = WorkbookFactory.create(inputStream);
        assertWorkbookEqual(workbook, expectWorkbook);
    }

    @Test
    public void testNullField() throws Exception {
        Object obj = new Object() {
            @ExcelMapping(columnIndex = 0)
            private final List<Object> nullList = null;
        };
        Workbook workbook = new ExcelExporter(obj)
                .generate()
                .getWorkbook();
    }

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

    @Ignore("test for performance")
    @Test
    public void test10000Row() throws Exception {
        Object source = new Object() {
            @ExcelMapping(columnIndex = 0)
            private final List<Bar> inners = getInners();

            private List<Bar> getInners() {
                List<Bar> list = new ArrayList<>(10000);
                for (int i = 0; i < 10000; i++) {
                    list.add(new Bar());
                }
                return list;
            }
        };
        URI uri = getClass().getClassLoader().getResource("test.xlsx").toURI();
        Path path = Paths.get(uri);
        path.toFile().createNewFile();
        new ExcelExporter(source)
                .generate()
                .output(new BufferedOutputStream(Files.newOutputStream(path)));
    }

    private static void assertWorkbookEqual(Workbook workbook, Workbook expectWorkbook) {
        Iterator<Sheet> expectSheetIterator = expectWorkbook.sheetIterator();
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while (expectSheetIterator.hasNext()) {
            Sheet expectSheet = expectSheetIterator.next();
            Sheet sheet = sheetIterator.next();
            Iterator<Row> expectRowIterator = expectSheet.rowIterator();
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (expectRowIterator.hasNext()) {
                Row expectRow = expectRowIterator.next();
                Row row = rowIterator.next();
                Iterator<Cell> expectCellIterator = expectRow.cellIterator();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (expectCellIterator.hasNext()) {
                    Cell expectCell = expectCellIterator.next();
                    Cell cell = cellIterator.next();
                    Assert.assertEquals(expectCell.toString(), cell.toString());
                    Assert.assertEquals(expectCell.getCellTypeEnum(), cell.getCellTypeEnum());
                }
            }
        }
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
        // 方便判等，所以写死
        @ExcelMapping(columnIndex = 5, tryFormatDateTime = true)
        private final long dateField = 1681873419533L;
        @ExcelMapping(sheetName = "iterable", columnIndex = 0)
        private final List<String> iterableField = Lists.newArrayList("1", "2", "3", "4");
        @ExcelRecursiveMapping
        private final Set<Bar> innerClassField = Sets.newHashSet(new Bar(), new Bar(), new Bar(), new Bar());
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

}