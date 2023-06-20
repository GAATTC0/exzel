package com.github.gaattc.exzel.excel;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author gaattc
 * @since 1.0
 * Created by gaattc on 2023/4/12
 */
@Slf4j
public class ExcelExporter {

    private final Object source;
    private Workbook workbook;

    public ExcelExporter(Object source) {
        this.source = source;
    }

    /**
     * 映射为excel工作簿对象
     */
    public ExcelExporter generate() throws Exception {
        if (null == workbook) {
            workbook = new ExcelGenerator(source).generate();
        }
        return this;
    }

    /**
     * 获取workbook
     */
    public Workbook getWorkbook() {
        if (null == workbook) {
            throw new IllegalStateException("workbook not generated, call generate() first");
        }
        return workbook;
    }

    /**
     * 将excel输出到HttpServletResponse
     *
     * @param response http返回值
     */
    public void response(HttpServletResponse response, String fileName) throws IOException {
        prepareResponse(response, fileName);
        ServletOutputStream outputStream = response.getOutputStream();
        flush(outputStream);
    }

    /**
     * 将excel输出到流
     *
     * @param stream 输出流
     */
    public void output(OutputStream stream) throws IOException {
        flush(stream);
    }

    public static void prepareResponse(HttpServletResponse response, String fileName) {
        response.setContentType("application/x-excel");
        response.setHeader("Content-disposition", "attachment; filename=" + fileName + ".xlsx");
    }

    private void flush(OutputStream outputStream) throws IOException {
        if (outputStream != null) {
            try {
                getWorkbook().write(outputStream);
                outputStream.flush();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            } finally {
                try {
                    ((SXSSFWorkbook) workbook).dispose();
                } catch (Throwable ignore) {
                }
                workbook.close();
            }
        }
    }

}
