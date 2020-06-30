package com.axpl.parsexls;

import com.aspose.cells.Cells;
import com.aspose.cells.License;
import com.aspose.cells.Range;
import com.aspose.cells.Worksheet;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.zip.ZipFile;

import static javax.xml.stream.XMLStreamConstants.END_ELEMENT;
import static javax.xml.stream.XMLStreamConstants.START_ELEMENT;

@Slf4j
public class ExcelConverter {

    private static final Map<Integer, String> staticColumns = new HashMap<>();

    static {
        for (int i = 0; i < 128; i++) {
            staticColumns.put(i, "column" + i);
        }
    }

    private SheetIterator iterator;

    public Iterator<Map<String, Object>> processFile(File file, Integer limit, String charset, ExcelType excelType) throws Exception {
        if (file == null ) {
            return null;
        }
        String fileName = file.getAbsolutePath();

        log.debug("file: {}", fileName);

        switch (excelType) {
            case ASPOSE:
                asposeAllExcel(file);
                break;
            case JEXCEL:
                oldFormatExcel(file, charset);
                break;
            case POI_XLS:
                xlsApachePoi(file);
                break;
            case POI_XLSX:
                xlsxApachePoi(file);
                break;
            case INCE:
                inceXlsxExcel(file);
                break;
        }

        return new Iterator<Map<String, Object>>() {
            int count;

            @Override
            public boolean hasNext() {
                return (limit == null || limit > count) && iterator.hasNext();
            }

            @Override
            public Map<String, Object> next() {
                count++;
                return iterator.next();
            }
        };

    }

    private void xlsxApachePoi(File file) throws IOException {
        log.warn("Use Apache POI library: {}", file.getAbsolutePath());
        Long start = System.currentTimeMillis();
        try (FileInputStream fileIn = new FileInputStream(file)) {
            final Workbook workbook = WorkbookFactory.create(fileIn);
            if (workbook.getNumberOfSheets() > 0) {
                final Sheet sheetAt = workbook.getSheetAt(0);
                log.info("Sheet name : {}", sheetAt.getSheetName());
            }
            iterator = new SheetIterator(workbook.getNumberOfSheets(),
                    (sheetIdx) -> getXlsIterator(workbook.getSheetAt(sheetIdx)),
                    (sheetIdx) -> workbook.getSheetAt(sheetIdx).getSheetName());
            iterator.setWorkBook(workbook);

            Long end = System.currentTimeMillis();
            final long l = end - start;
            log.info("xlsxApachePoiExcel: {}", l);
        }
    }

    private void xlsApachePoi(File file) throws IOException {
        try (InputStream inputStream = new FileInputStream(file)) {
            Workbook workbook = new HSSFWorkbook(inputStream, false);
            iterator = new SheetIterator(workbook.getNumberOfSheets(),
                    (sheetIdx) -> getXlsIterator(workbook.getSheetAt(sheetIdx)),
                    (sheetIdx) -> workbook.getSheetAt(sheetIdx).getSheetName());
            iterator.setWorkBook(workbook);
        }
    }

    private void inceXlsxExcel(File file) {
        log.warn("Use Ince-excel library: {}", file.getAbsolutePath());
        Long start = System.currentTimeMillis();
        SimpleXLSXWorkbook workbook = new SimpleXLSXWorkbook(file);
        Long end = System.currentTimeMillis();
        final long l = end - start;
        log.info("office2007Iterator: {}", l);
        List<String> names = parseSheetNames(file);
        if (workbook.getSheetCount() > 0) {
            final com.incesoft.tools.excel.xlsx.Sheet sheet = workbook.getSheet(0, false);
            log.info("Sheet 0 size: {}", sheet.getRowCount());
        }

        iterator = new SheetIterator(workbook.getSheetCount(),
                (sheetIdx) -> getXlsxIterator(workbook.getSheet(sheetIdx, false)),
                (sheetIdx) -> names.size() > sheetIdx ? names.get(sheetIdx) : "" + sheetIdx);
        iterator.setWorkBook(workbook);


    }

    public static List<String> parseSheetNames(File file) {
        List<String> names = new ArrayList<>();
        try (ZipFile zf = new ZipFile(file)) {
            XMLInputFactory xmlif = XMLInputFactory.newFactory();
            try (InputStream in = zf.getInputStream(zf.getEntry("xl/workbook.xml"))) {
                XMLStreamReader xmlr = xmlif.createXMLStreamReader(in);
                while (xmlr.hasNext()) {
                    int status = xmlr.nextTag();
                    if (END_ELEMENT == status && "sheets".equals(xmlr.getName().getLocalPart())) {
                        break;
                    }
                    if (START_ELEMENT == status && "sheet".equals(xmlr.getName().getLocalPart())) {
                        for (int i = 0; i < xmlr.getAttributeCount(); i++) {
                            if ("name".equals(xmlr.getAttributeName(i).getLocalPart())) {
                                names.add(xmlr.getAttributeValue(i));
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.warn("Can't detect sheets names: " + file.getAbsolutePath(), e);
        }
        return names;
    }

    private void oldFormatExcel(File file, String charset) throws IOException, BiffException {
        log.warn("Use Jexcel API library: {}", file.getAbsolutePath());
        WorkbookSettings ws = new WorkbookSettings();
        ws.setEncoding(charset);
        ws.setWriteAccess("n/a");
        jxl.Workbook workbook = jxl.Workbook.getWorkbook(file, ws);
        iterator = new SheetIterator(workbook.getNumberOfSheets(),
                (sheetIdx) -> getXlsIterator(workbook.getSheet(sheetIdx)),
                (sheetIdx) -> workbook.getSheet(sheetIdx).getName());
        iterator.setWorkBook(workbook);
    }

    public void closeIterator() {
        iterator.closeWorkBook();
    }



    private void asposeAllExcel(File file) throws Exception {
        log.warn("Use Aspose library: {}", file.getAbsolutePath());
        getAsposeLicense();
        try (FileInputStream fis = new FileInputStream(file)) {
            final com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(fis);
            if (workbook.getWorksheets().getCount() > 0) {
                iterator = new SheetIterator(workbook.getWorksheets().getCount(),
                        (sheetIdx) -> getXlsIterator(workbook.getWorksheets().get(sheetIdx)),
                        (sheetIdx) -> workbook.getWorksheets().get(sheetIdx).getName());
                iterator.setWorkBook(workbook);
            }
        }
    }

    public void getAsposeLicense() throws Exception {
        try(InputStream license = ExcelConverter.class.getResourceAsStream("license.xml")) {
            License aposeLic = new License();
            aposeLic.setLicense(license);
        } catch (Exception e) {
            log.error("Can't apply license: ", e);
            throw new Exception(e);
        }
    }

    public static String getStaticColumn(int index) {
        return staticColumns.computeIfAbsent(index, e -> "column" + index);
    }

    private Iterator<Map<String, Object>> getXlsIterator(Worksheet sheet) {
        log.debug("Sheet: {}", sheet.getName());
        return new Iterator<Map<String, Object>>() {
            int count = 0;

            @Override
            public boolean hasNext() {
                if (sheet.getCells() == null || sheet.getCells().getMaxDisplayRange() == null) {
                    return false;
                }
                int rowCount = sheet.getCells().getMaxDisplayRange().getRowCount();
                return (count < rowCount) ;
            }

            @Override
            public Map<String, Object> next() {
                Map<String, Object> mapRow = new LinkedHashMap<>();
                Range maxDisplayRange = sheet.getCells().getMaxDisplayRange();
                int columnCount = maxDisplayRange.getColumnCount();

                mapRow.put("current_row", "" + count);
                Cells cells = sheet.getCells();

                for (int column = 0; column < columnCount; column++) {
                    com.aspose.cells.Cell cell = cells.get(count, column);
                    String key = getStaticColumn(column);
                    if (cell == null) {
                        mapRow.put(key, null);
                        continue;
                    }
                    String strValue;

                    switch (cell.getType()) {
                        case com.aspose.cells.CellValueType.IS_BOOL:
                            mapRow.put(key, "" + cell.getValue());
                            break;
                        case com.aspose.cells.CellValueType.IS_DATE_TIME:
                            mapRow.put(key, "" + cell.getValue());
                            break;
                        case com.aspose.cells.CellValueType.IS_NUMERIC:
                            strValue = "" + cell.getValue();
                            if (strValue.split(",").length > 3) {
                                strValue = strValue.replace(",", "");
                            } else if (strValue.indexOf(',') + 3 < strValue.length()) {
                                strValue = strValue.replace(",", "");
                            } else if (strValue.indexOf(',') > 0 && strValue.indexOf('.') > 0) {
                                if (strValue.indexOf(',') > strValue.indexOf('.')) {
                                    strValue = strValue.replace(".", "");
                                } else {
                                    strValue = strValue.replace(",", "");
                                }
                            }
                            mapRow.put(key, strValue);
                            break;
                        case com.aspose.cells.CellValueType.IS_STRING:
                            mapRow.put(key, "" + cell.getValue());
                            break;
                        case com.aspose.cells.CellValueType.IS_NULL:
                            mapRow.put(key, "");
                            break;
                        default:
                            mapRow.put(key, "" + cell.getValue());
                    }
                }
                if (log.isDebugEnabled() && count == 0) {
                    log.debug("First row: {}", mapRow);
                }
                count++;
                return mapRow;
            }
        };
    }

    private Iterator<Map<String, Object>> getXlsIterator(jxl.Sheet sheet) {
        log.debug("Sheet: {}", sheet.getName());
        return new Iterator<Map<String, Object>>() {
            int count = 0;

            @Override
            public boolean hasNext() {
                return (count < sheet.getRows());
            }

            @Override
            public Map<String, Object> next() {
                Map<String, Object> mapRow = new LinkedHashMap<>();
                mapRow.put("current_row", "" + count);
                jxl.Cell[] row = sheet.getRow(count);
                for (int cn = 0; cn < row.length; cn++) {
                    jxl.Cell cell = row[cn];
                    String key = getStaticColumn(cn);
                    if (cell == null) {
                        mapRow.put(key, null);
                        continue;
                    }
                    String strValue;
                    final jxl.CellType type = cell.getType();
                    if (type == jxl.CellType.EMPTY) {
                        mapRow.put(key, "");
                    } else if (type == jxl.CellType.BOOLEAN) {
                        strValue = "" + cell.getContents();
                        mapRow.put(key, strValue);
                    } else if (type == jxl.CellType.ERROR) {
                        mapRow.put(key, "#err");
                    } else if (type == jxl.CellType.NUMBER) {
                        strValue = "" + cell.getContents();
                        if (strValue.split(",").length > 3) {
                            strValue = strValue.replace(",", "");
                        } else if (strValue.indexOf(',') + 3 < strValue.length()) {
                            strValue = strValue.replace(",", "");
                        } else if (strValue.indexOf(',') > 0 && strValue.indexOf('.') > 0) {
                            if (strValue.indexOf(',') > strValue.indexOf('.')) {
                                strValue = strValue.replace(".", "");
                            } else {
                                strValue = strValue.replace(",", "");
                            }
                        }
                        mapRow.put(key, strValue);
                    } else {
                        strValue = "" + cell.getContents();
                        mapRow.put(key, strValue);
                    }
                }
                if (log.isDebugEnabled() && count == 0) {
                    log.debug("First row: {}", mapRow);
                }
                count++;
                return mapRow;
            }
        };
    }

    private Iterator<Map<String, Object>> getXlsIterator(Sheet sheet) {
        Iterator<Row> rit = sheet.rowIterator();
        if (!rit.hasNext()) {
            return null;
        }
        log.debug("Sheet: {}", sheet.getSheetName());
        return new Iterator<Map<String, Object>>() {
            int count = 0;

            @Override
            public boolean hasNext() {
                return (count == 0 || rit.hasNext()) ;
            }

            @Override
            public Map<String, Object> next() {
                Map<String, Object> mapRow = new LinkedHashMap<>();
                mapRow.put("current_row", "" + count);
                Row row = rit.next();
                for (int cn = 0; cn < row.getLastCellNum(); cn++) {
                    Cell cell = row.getCell(cn, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    String key = getStaticColumn(cn);
                    if (cell == null) {
                        mapRow.put(key, null);
                        continue;
                    }
                    String strValue;
                    switch (cell.getCellType()) {
                        case BLANK:
                            mapRow.put(key, "");
                            break;
                        case BOOLEAN:
                            strValue = cell.getBooleanCellValue() ? "true" : "false";
                            mapRow.put(key, strValue);
                            break;
                        case ERROR:
                            mapRow.put(key, "#err");
                            break;
                        case FORMULA:
                            FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
                            CellValue cellValue = evaluator.evaluate(cell);
                            if (cellValue.getStringValue() == null) {
                                mapRow.put(key, "" + cellValue.getNumberValue());
                            } else {
                                mapRow.put(key, cellValue.getStringValue());
                            }
                            break;
                        case NUMERIC:
                            double value = cell.getNumericCellValue();
                            if (value % 1 == 0) {
                                strValue = "" + ((long) value);
                                mapRow.put(key, strValue);
                            } else {
                                strValue = "" + value;
                                mapRow.put(key, strValue);
                            }
                            break;
                        case STRING:
                            strValue = cell.getRichStringCellValue().getString();
                            mapRow.put(key, strValue);
                            break;
                        default:
                            throw new RuntimeException("Unknown cell type: " + cell.getCellType() + " for " + cell);
                    }

                }
                if (log.isDebugEnabled() && count == 0) {
                    log.debug("First row: {}", mapRow);
                }

                count++;
                return mapRow;
            }
        };
    }

    @SneakyThrows
    public Iterator<Map<String, Object>> getXlsxIterator(com.incesoft.tools.excel.xlsx.Sheet sheet) {
        // workaround for bug in library: duplicates cells
        com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader reader = sheet.newReader();
        Class cls = com.incesoft.tools.excel.xlsx.Cell.class;
        Field field = cls.getDeclaredField("r");
        field.setAccessible(true);
        return new Iterator<Map<String, Object>>() {
            int count = 0;
            com.incesoft.tools.excel.xlsx.Cell[] row = reader.readRow();

            @Override
            public boolean hasNext() {
                return row != null;
            }

            @Override
            @SneakyThrows
            public Map<String, Object> next() {

                Map<String, Object> mapRow = new LinkedHashMap<>();
                mapRow.put("current_row", "" + count);
                int colNum = 0;
                Set<String> processed = new HashSet<>();
                for (com.incesoft.tools.excel.xlsx.Cell cell : row) {
                    String key = getStaticColumn(colNum);
                    if (cell == null) {
                        colNum++;
                        mapRow.put(key, null);
                        continue;
                    }
                    String cellName = (String) field.get(cell);
                    if (processed.contains(cellName)) {
                        colNum++;
                        mapRow.put(key, null);
                        continue;
                    }
                    processed.add(cellName);
                    String strValue = cell.getValue();
                    mapRow.put(key, strValue);
                    colNum++;
                }
                if (log.isDebugEnabled() && count == 0) {
                    log.debug("First row: {}", mapRow);
                }

                count++;
                row = reader.readRow();
                return mapRow;
            }
        };
    }
}
