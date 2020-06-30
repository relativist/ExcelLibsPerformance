package com.axpl.parsexls;

import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;
import jxl.read.biff.WorkbookParser;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.Closeable;
import java.util.Collections;
import java.util.Iterator;
import java.util.Map;
import java.util.function.Function;

@Slf4j
public class SheetIterator implements Iterator<Map<String, Object>> {

    private int sheetCount;
    private int current;
    private Function<Integer, Iterator<Map<String, Object>>> sheetIteratorProducer;
    private Iterator<Map<String, Object>> currentIterator;

    private Function<Integer, String> sheetNameFunc;
    private String sheetName;
    private Object workBook;

    SheetIterator(int sheetCount, Function<Integer, Iterator<Map<String, Object>>> sheetIteratorProducer, Function<Integer, String> sheetNameFunc) {
        this.sheetCount = sheetCount;
        this.sheetIteratorProducer = sheetIteratorProducer;
        currentIterator = sheetIteratorProducer.apply(0);
        this.sheetNameFunc = sheetNameFunc;
        this.sheetName = sheetNameFunc.apply(0);
    }

    void setWorkBook(Object workBook) {
        this.workBook = workBook;
    }

    void closeWorkBook() {
        log.info("Close workbook: {}", workBook == null ? "null" : workBook.getClass());
        if (workBook != null) {
            try {
                if (workBook.getClass().isAssignableFrom(HSSFWorkbook.class)) {
                    log.info("Close HSSFWorkbook.");
                    ((Workbook) workBook).close();
                } else if (workBook.getClass().isAssignableFrom(jxl.Workbook.class)) {
                    log.info("Close jxl.Workbook.");
                    ((jxl.Workbook) workBook).close();
                } else if (workBook.getClass().isAssignableFrom(SimpleXLSXWorkbook.class)) {
                    log.info("Close SimpleXLSXWorkbook.");
                    ((SimpleXLSXWorkbook) workBook).close();
                } else if (workBook.getClass().isAssignableFrom(com.aspose.cells.Workbook.class)) {
                    log.info("Close aspose.");
                    ((com.aspose.cells.Workbook) workBook).dispose();
                } else if (workBook.getClass().isAssignableFrom(Closeable.class)) {
                    log.info("Close closable.");
                    ((Closeable) workBook).close();
                } else if (workBook.getClass().isAssignableFrom(WorkbookParser.class)) {
                    log.info("Close WorkbookParser.");
                    ((WorkbookParser) workBook).close();
                } else if (workBook.getClass().isAssignableFrom(XSSFWorkbook.class)) {
                    log.info("Close XSSFWorkbook.");
                    ((XSSFWorkbook) workBook).close();
                } else {
                    log.error("Unknown workbook type: {}", workBook.getClass().getSimpleName());
                }
            } catch (Exception e) {
                log.error("Error on close workbook {}", e.getMessage());
            }
        } else {
            log.info("Workbook is null (already closed)");
        }
    }

    @Override
    public boolean hasNext() {
        if (currentIterator == null) {
            return false;
        }
        boolean rs = currentIterator.hasNext();
        if (rs) {
            return true;
        }
        current++;
        if (current < sheetCount) {
            currentIterator = sheetIteratorProducer.apply(current);
            sheetName = sheetNameFunc.apply(current);
            rs = hasNext();
        } else {
            rs = false;
        }
        return rs;
    }

    @Override
    public Map<String, Object> next() {
        return currentIterator.next();
    }
}
