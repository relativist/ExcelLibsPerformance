package com.axpl.parsexls.dto;

import com.axpl.parsexls.ExcelType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ProcessExcelResult {
    private ExcelType type;
    private Long countLines;
    private Long durationMs;
    private String fileName;

    @Override
    public String toString() {
        return "Type=" + type + ", CountLines=" + countLines + ", FileName=" + fileName + ", DurationMs=" + durationMs ;
    }
}
