package com.axpl.parsexls;

import com.axpl.parsexls.dto.ProcessExcelResult;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.BeanFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Service
public class ParserManager {

    @SneakyThrows
    public void run() {
        log.info("Start parse files....");

        List<ProcessExcelResult> results = new ArrayList<>();
        List<String> fileNames = new ArrayList<>();
        fileNames.add("small.xls");
        fileNames.add("big.xls");
        fileNames.add("big.xlsx");

        for (String fileName : fileNames) {
            final String lowerCase = fileName.toLowerCase();
            if (lowerCase.endsWith("xls")) {
                results.add(getDuration(ExcelType.JEXCEL, fileName));
                results.add(getDuration(ExcelType.ASPOSE, fileName));
                results.add(getDuration(ExcelType.POI_XLS, fileName));
            } else if (lowerCase.endsWith("xlsx")) {
                results.add(getDuration(ExcelType.INCE, fileName));
                results.add(getDuration(ExcelType.ASPOSE, fileName));
                results.add(getDuration(ExcelType.POI_XLSX, fileName));
            }
        }

        String version = System.getProperty("java.version");
        System.out.println("JAVA VERSION: " + version);
        System.out.println("--------------[RESULT]----------------");
        results.stream()
                .collect(Collectors.groupingBy(ProcessExcelResult::getFileName))
                .forEach((key, value) -> {

                    value.sort(Comparator.comparing(ProcessExcelResult::getDurationMs));
                    System.out.println("File: " + key);
                    value.forEach(System.out::println);
                    System.out.println("-------------------------");
                });
    }

    private ProcessExcelResult getDuration(ExcelType excelType, String fileName) throws Exception {
        File xlsFile = new File(fileName);
        if (!xlsFile.exists()) {
            log.info("File not exists: {}", fileName);
            return new ProcessExcelResult(excelType, 0L, 0L, fileName);
        }
        final long start = System.currentTimeMillis();
        ExcelConverter converter = new ExcelConverter();
        final Iterator<Map<String, Object>> iterator = converter.processFile(xlsFile, null, "UTF-8", excelType);
        final long lineCount = processIterator(iterator);
        converter.closeIterator();
        final long end = System.currentTimeMillis();
        final long duration = end - start;
        log.info("=========>Type: {} / File: {} / lines: {} Duration: {}", excelType, fileName, lineCount, duration);
        return new ProcessExcelResult(excelType, lineCount, duration, fileName);
    }

    private long processIterator(Iterator<Map<String, Object>> iterator) {
        long lineCount = 0;
        while (iterator.hasNext()) {
            final Map<String, Object> next = iterator.next();
            if (next.size() == 0) {
                log.info("empty line");
            }
            lineCount++;
        }
        log.info("Read lines: {}", lineCount);
        return lineCount;
    }


    @PostConstruct
    public void destroy() {
        log.info("SHUT DOWN....");
    }


}
