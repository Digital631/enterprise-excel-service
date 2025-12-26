package com.sgcc.easyexcel.service.impl;

import com.sgcc.easyexcel.config.ExcelConfig;
import com.sgcc.easyexcel.dto.request.MultiSheetExportRequest;
import com.sgcc.easyexcel.dto.request.SheetDataConfig;
import com.sgcc.easyexcel.dto.request.SingleSheetExportRequest;
import com.sgcc.easyexcel.dto.response.ExcelFileInfo;
import com.sgcc.easyexcel.enums.ExcelErrorCode;
import com.sgcc.easyexcel.exception.ExcelExportException;
import com.sgcc.easyexcel.service.ExcelExportService;
import com.sgcc.easyexcel.util.ExcelTemplateUtil;
import com.sgcc.easyexcel.util.FileNameGenerator;
import com.sgcc.easyexcel.util.TaskManager;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.stream.Collectors;

/**
 * Excel导出服务实现
 * 提供单Sheet和多Sheet的Excel导出功能
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Service
public class ExcelExportServiceImpl implements ExcelExportService {

    private final ExcelConfig excelConfig;
    private final FileNameGenerator fileNameGenerator;
    private final TaskManager taskManager;
    private final ExcelTemplateUtil excelTemplateUtil;

    public ExcelExportServiceImpl(ExcelConfig excelConfig, 
                                  FileNameGenerator fileNameGenerator,
                                  TaskManager taskManager,
                                  ExcelTemplateUtil excelTemplateUtil) {
        this.excelConfig = excelConfig;
        this.fileNameGenerator = fileNameGenerator;
        this.taskManager = taskManager;
        this.excelTemplateUtil = excelTemplateUtil;
    }

    @Override
    public ExcelFileInfo exportSingleSheet(SingleSheetExportRequest request) {
        return exportSingleSheet(request, null);
    }

    @Override
    public ExcelFileInfo exportSingleSheet(SingleSheetExportRequest request, String taskId) {
        long startTime = System.currentTimeMillis();
        log.info("开始单Sheet导出：template={}, dataSize={}, bigDataMode={}", 
                request.getTemplatePath(), 
                request.getData() != null ? request.getData().size() : 0,
                request.getEnableBigDataMode());

        try {
            // 获取模板文件
            String templatePath = getTemplatePath(request.getTemplatePath());
            File templateFile = new File(templatePath);
            if (!templateFile.exists()) {
                log.error("模板文件不存在：{}", templatePath);
                throw new ExcelExportException(ExcelErrorCode.TEMPLATE_NOT_FOUND, 
                    "模板文件不存在：" + templatePath);
            }

            // 生成输出文件路径
            String outputPath = generateOutputPath(request.getTemplatePath(), request);
                    
            // 处理文件重名问题（确保duplicateStrategy配置为rename）
            outputPath = fileNameGenerator.handleDuplicateFile(outputPath);
                    
            File outputFile = new File(outputPath);

            // 检查是否启用大数据模式
            if (request.getEnableBigDataMode() && 
                request.getData() != null && 
                request.getData().size() > excelConfig.getBigDataThreshold()) {
                
                log.info("启用大数据模式导出：threshold={}, actualSize={}", 
                    excelConfig.getBigDataThreshold(), request.getData().size());
                
                ExcelFileInfo result = exportLargeDataWithMultipleSheets(templateFile, outputFile, request, taskId);
                
                if (taskId != null) {
                    taskManager.setTaskResult(taskId, result, TaskManager.TaskStatus.COMPLETED);
                }
                
                // 设置导出耗时
                long duration = System.currentTimeMillis() - startTime;
                result.setExportDuration(duration);
                result.setExportDurationSeconds(BigDecimal.valueOf(duration / 1000.0).setScale(2, BigDecimal.ROUND_HALF_UP));
                
                return result;
            } else {
                log.info("使用普通模式导出");
                
                ExcelFileInfo result = exportNormalMode(templateFile, outputFile, request);
                
                if (taskId != null) {
                    taskManager.setTaskResult(taskId, result, TaskManager.TaskStatus.COMPLETED);
                }
                
                // 设置导出耗时
                long duration = System.currentTimeMillis() - startTime;
                result.setExportDuration(duration);
                result.setExportDurationSeconds(BigDecimal.valueOf(duration / 1000.0).setScale(2, BigDecimal.ROUND_HALF_UP));
                
                return result;
            }
        } catch (InterruptedException e) {
            log.info("导出任务被中断：taskId={}", taskId);
            if (taskId != null) {
                taskManager.setTaskResult(taskId, "导出任务被中断", TaskManager.TaskStatus.STOPPED);
            }
            throw new ExcelExportException(ExcelErrorCode.TASK_INTERRUPTED, "导出操作被中断");
        } catch (Exception e) {
            log.error("单Sheet导出失败：taskId={}", taskId, e);
            if (taskId != null) {
                taskManager.setTaskResult(taskId, e.getMessage(), TaskManager.TaskStatus.FAILED);
            }
            if (e instanceof ExcelExportException) {
                throw new ExcelExportException(ExcelErrorCode.UNKNOWN_ERROR,
                        e.getMessage());
            } else {
                throw new ExcelExportException(ExcelErrorCode.FILE_WRITE_ERROR, 
                    "文件写入失败：" + e.getMessage());
            }
        }
    }

    @Override
    public ExcelFileInfo exportMultiSheet(MultiSheetExportRequest request) {
        long startTime = System.currentTimeMillis();
        log.info("开始多Sheet导出：template={}, sheetCount={}", 
                request.getTemplatePath(), 
                request.getSheetDataList() != null ? request.getSheetDataList().size() : 0);

        try {
            // 获取模板文件
            String templatePath = getTemplatePath(request.getTemplatePath());
            File templateFile = new File(templatePath);
            if (!templateFile.exists()) {
                log.error("模板文件不存在：{}", templatePath);
                throw new ExcelExportException(ExcelErrorCode.TEMPLATE_NOT_FOUND, 
                    "模板文件不存在：" + templatePath);
            }

            // 生成输出文件路径（使用默认配置）
            String outputPath = generateOutputPath(request.getTemplatePath(), null);
                    
            // 处理文件重名问题（确保duplicateStrategy配置为rename）
            outputPath = fileNameGenerator.handleDuplicateFile(outputPath);
                    
            File outputFile = new File(outputPath);

            // 如果有占位符数据，使用EasyExcel模板功能处理
            if (request.getPlaceholders() != null && !request.getPlaceholders().isEmpty()) {
                // 先用占位符填充模板
                excelTemplateUtil.fillTemplate(templateFile.getAbsolutePath(), 
                    outputFile.getAbsolutePath(), 
                    null, // 多Sheet导出时，数据是通过SheetDataConfig分批处理的，不直接传递数据列表
                    request.getPlaceholders());
                
                // 然后使用填充后的文件作为基础，添加多Sheet数据
                addMultiSheetData(outputFile, request);
            } else {
                // 使用SXSSFWorkbook进行流式处理，优化内存使用
                try (FileInputStream fis = new FileInputStream(templateFile);
                     XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis)) {

                    // 创建流式工作簿
                    try (SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 1)) {
                        
                        // 先删除模板中所有的Sheet
                        int numberOfSheets = sxssfWorkbook.getNumberOfSheets();
                        for (int i = numberOfSheets - 1; i >= 0; i--) {
                            sxssfWorkbook.removeSheetAt(i);
                        }

                        // 遍历每个Sheet配置
                        for (int i = 0; i < request.getSheetDataList().size(); i++) {
                            SheetDataConfig sheetConfig = request.getSheetDataList().get(i);
                            
                            // 创建或获取Sheet
                            String sheetName = sheetConfig.getSheetName() != null ? sheetConfig.getSheetName() : "Sheet" + (i + 1);
                            Sheet sheet = sxssfWorkbook.createSheet(sheetName);

                            // 填充数据到Sheet
                            fillDataToSheet(sheet, sheetConfig.getData(), sheetConfig.getStartRow(), 
                                sheetConfig.getFieldMapping());
                        }

                        // 写入文件
                        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                            sxssfWorkbook.write(fos);
                        }
                        
                        // 清理临时文件
                        sxssfWorkbook.dispose();
                    }
                }
            }

            // 返回文件信息
            ExcelFileInfo fileInfo = new ExcelFileInfo();
            fileInfo.setFilePath(outputFile.getAbsolutePath());
            fileInfo.setFileName(outputFile.getName());
            fileInfo.setFileSize(outputFile.length());
            fileInfo.setExportTime(LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
            
            // 设置导出耗时
            long duration = System.currentTimeMillis() - startTime;
            fileInfo.setExportDuration(duration);
            fileInfo.setExportDurationSeconds(BigDecimal.valueOf(duration / 1000.0).setScale(2, BigDecimal.ROUND_HALF_UP));
            
            log.info("多Sheet导出完成：file={}, size={}bytes, 耗时: {}ms", 
                outputFile.getName(), outputFile.length(), duration);
            return fileInfo;

        } catch (Exception e) {
            log.error("多Sheet导出失败", e);
            throw new ExcelExportException(ExcelErrorCode.FILE_WRITE_ERROR, 
                "文件写入失败：" + e.getMessage());
        }
    }

    /**
     * 普通模式导出
     */
    private ExcelFileInfo exportNormalMode(File templateFile, File outputFile, SingleSheetExportRequest request) 
            throws Exception {
        long startTime = System.currentTimeMillis();
        
        // 检查是否同时有占位符和数据填充需求
        if ((request.getPlaceholders() != null && !request.getPlaceholders().isEmpty()) && 
            (request.getData() != null && !request.getData().isEmpty())) {
            // 先使用EasyExcel模板功能填充占位符
            excelTemplateUtil.fillTemplate(templateFile.getAbsolutePath(), 
                outputFile.getAbsolutePath(), 
                null, // 先不填充数据，只填充占位符
                request.getPlaceholders());
            
            // 然后使用POI直接填充数据到指定位置
            fillDataToExcelFile(outputFile, request);
        } else if (request.getPlaceholders() != null && !request.getPlaceholders().isEmpty()) {
            // 只有占位符，使用EasyExcel模板功能
            excelTemplateUtil.fillTemplate(templateFile.getAbsolutePath(), 
                outputFile.getAbsolutePath(), 
                null, 
                request.getPlaceholders());
            
            // 如果需要更新Sheet名称，单独处理
            if (request.getSheetName() != null && !request.getSheetName().isEmpty()) {
                updateSheetName(outputFile, request.getSheetName());
            }
        } else if (request.getData() != null && !request.getData().isEmpty()) {
            // 检查数据中是否包含文档级占位符（只有一行数据时）
            if (request.getData().size() == 1 && hasDocumentLevelPlaceholders(templateFile, request.getData().get(0))) {
                // 数据中包含文档级占位符，合并到placeholders中
                Map<String, Object> placeholders = new HashMap<>(request.getPlaceholders() != null ? request.getPlaceholders() : new HashMap<>());
                placeholders.putAll(request.getData().get(0));
                
                excelTemplateUtil.fillTemplate(templateFile.getAbsolutePath(), 
                    outputFile.getAbsolutePath(), 
                    null, 
                    placeholders);
                
                // 如果需要更新Sheet名称，单独处理
                if (request.getSheetName() != null && !request.getSheetName().isEmpty()) {
                    updateSheetName(outputFile, request.getSheetName());
                }
            } else {
                // 只有数据，使用POI直接填充
                fillDataToExcelFile(outputFile, request);
            }
        } else {
            // 既没有占位符也没有数据，直接复制模板
            try (FileInputStream fis = new FileInputStream(templateFile);
                 FileOutputStream fos = new FileOutputStream(outputFile)) {
                fis.transferTo(fos);
            }
            
            // 如果需要更新Sheet名称，单独处理
            if (request.getSheetName() != null && !request.getSheetName().isEmpty()) {
                updateSheetName(outputFile, request.getSheetName());
            }
        }

        // 返回文件信息
        ExcelFileInfo fileInfo = new ExcelFileInfo();
        fileInfo.setFilePath(outputFile.getAbsolutePath());
        fileInfo.setFileName(outputFile.getName());
        fileInfo.setFileSize(outputFile.length());
        fileInfo.setExportTime(LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
        
        // 设置导出耗时
        long duration = System.currentTimeMillis() - startTime;
        fileInfo.setExportDuration(duration);
        fileInfo.setExportDurationSeconds(BigDecimal.valueOf(duration / 1000.0).setScale(2, BigDecimal.ROUND_HALF_UP));
        
        log.info("普通模式导出完成：file={}, size={}bytes, 耗时: {}ms", 
            outputFile.getName(), outputFile.length(), duration);
        return fileInfo;
    }

    /**
     * 大数据模式导出（多Sheet）
     */
    private ExcelFileInfo exportLargeDataWithMultipleSheets(File templateFile, File outputFile, 
            SingleSheetExportRequest request, String taskId) throws Exception {
        long startTime = System.currentTimeMillis();
        
        List<Map<String, Object>> allData = request.getData();
        int batchSize = request.getBatchSize() > 0 ? request.getBatchSize() : excelConfig.getDefaultBatchSize();
        int totalSize = allData.size();
        
        log.info("开始大数据模式导出：totalSize={}, batchSize={}", totalSize, batchSize);

        // 使用SXSSFWorkbook进行流式处理，优化内存使用
        try (FileInputStream fis = new FileInputStream(templateFile);
             XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis)) {

            // 创建流式工作簿，内存中只保留1行，其余写入磁盘
            try (SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 1)) {
                
                // 先删除模板中所有的Sheet
                int numberOfSheets = sxssfWorkbook.getNumberOfSheets();
                for (int i = numberOfSheets - 1; i >= 0; i--) {
                    sxssfWorkbook.removeSheetAt(i);
                }
                
                int sheetIndex = 0;
                int processedCount = 0;
                
                // 分批处理数据
                for (int i = 0; i < totalSize; i += batchSize) {
                    // 检查线程是否被中断
                    if (taskId != null && Thread.currentThread().isInterrupted()) {
                        log.info("导出任务被中断：taskId={}", taskId);
                        throw new InterruptedException("导出任务被中断");
                    }

                    int endIndex = Math.min(i + batchSize, totalSize);
                    List<Map<String, Object>> batchData = allData.subList(i, endIndex);
                    
                    String sheetName = (request.getSheetName() != null ? request.getSheetName() : "Sheet") + (sheetIndex + 1);
                    
                    // 创建新Sheet
                    Sheet currentSheet = sxssfWorkbook.createSheet(sheetName);

                    // 从模板复制表头（这里简化处理，实际可能需要更复杂的模板处理）
                    createHeaderRow(currentSheet, request.getFieldMapping(), request.getStartRow());

                    // 填充数据到当前Sheet
                    fillDataToSheet(currentSheet, batchData, request.getStartRow(), request.getFieldMapping());
                    
                    sheetIndex++;
                    
                    // 更新进度
                    processedCount += batchData.size();
                    
                    // 计算并显示进度
                    int progress = (int) ((double) processedCount / totalSize * 100);
                    log.info("完成Sheet {} 导出：dataCount={}, 总进度: {}% ({}/{})", 
                        sheetName, batchData.size(), progress, processedCount, totalSize);
                }

                // 写入文件
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    sxssfWorkbook.write(fos);
                }
                
                // 在完成写入后，清理临时文件
                sxssfWorkbook.dispose();
            }
        }

        // 返回文件信息
        ExcelFileInfo fileInfo = new ExcelFileInfo();
        fileInfo.setFilePath(outputFile.getAbsolutePath());
        fileInfo.setFileName(outputFile.getName());
        fileInfo.setFileSize(outputFile.length());
        fileInfo.setExportTime(LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
        
        // 设置导出耗时
        long duration = System.currentTimeMillis() - startTime;
        fileInfo.setExportDuration(duration);
        fileInfo.setExportDurationSeconds(BigDecimal.valueOf(duration / 1000.0).setScale(2, BigDecimal.ROUND_HALF_UP));
        
        log.info("大数据模式导出完成：file={}, size={}bytes, 耗时: {}ms", 
            outputFile.getName(), outputFile.length(), duration);
        return fileInfo;
    }

    /**
     * 创建表头行
     */
    private void createHeaderRow(Sheet sheet, Map<String, String> fieldMapping, int startRow) {
        if (fieldMapping != null && startRow > 1) {
            Row headerRow = sheet.createRow(startRow - 2); // 表头通常在数据行之前
            int colIndex = 0;
            for (String header : fieldMapping.keySet()) {
                Cell cell = headerRow.createCell(colIndex);
                cell.setCellValue(header);
                colIndex++;
            }
        }
    }

    /**
     * 应用模板样式到数据行
     */
    private void applyTemplateStyles(Sheet sheet, Map<Integer, Map<Integer, CellStyle>> templateStyles, 
            int startRow, int dataRowCount) {
        // 应用模板样式到数据行
        for (int i = 0; i < dataRowCount; i++) {
            int templateRowNum = startRow - 1; // 假设模板数据行样式在起始行前一行
            if (templateStyles.containsKey(templateRowNum)) {
                Map<Integer, CellStyle> templateRowStyles = templateStyles.get(templateRowNum);
                
                Row dataRow = sheet.getRow(startRow + i);
                if (dataRow == null) {
                    dataRow = sheet.createRow(startRow + i);
                }
                
                for (Map.Entry<Integer, CellStyle> entry : templateRowStyles.entrySet()) {
                    int colIndex = entry.getKey();
                    CellStyle style = entry.getValue();
                    
                    Cell cell = dataRow.getCell(colIndex);
                    if (cell == null) {
                        cell = dataRow.createCell(colIndex);
                    }
                    cell.setCellStyle(style);
                }
            }
        }
    }

    /**
     * 填充数据到Sheet
     */
    private void fillDataToSheet(Sheet sheet, List<Map<String, Object>> data, 
            int startRow, Map<String, String> fieldMapping) {
        
        if (data == null || data.isEmpty()) {
            return;
        }

        // 使用提供的字段映射或从模板第一行获取表头
        Map<String, Integer> headerColumnMap = new HashMap<>();
        if (fieldMapping != null && !fieldMapping.isEmpty()) {
            // 使用提供的字段映射
            for (Map.Entry<String, String> entry : fieldMapping.entrySet()) {
                String templateField = entry.getKey(); // 模板中的字段名，如"姓名"
                int columnIndex = findColumnIndex(sheet, templateField);
                if (columnIndex != -1) {
                    headerColumnMap.put(templateField, columnIndex);
                }
            }
        } else {
            // 如果没有提供字段映射，尝试从模板第一行获取表头
            Row headerRow = sheet.getRow(0); // 假设标题行在第0行
            if (headerRow != null) {
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    Cell cell = headerRow.getCell(i);
                    if (cell != null) {
                        String headerName = getCellValueAsString(cell);
                        if (headerName != null && !headerName.trim().isEmpty()) {
                            headerColumnMap.put(headerName, i);
                        }
                    }
                }
            }
        }

        // 按照原有逻辑，在指定行开始填充数据
        // 注意：startRow是从1开始的，但在POI中行索引从0开始，所以实际索引为startRow - 1
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.getRow(startRow - 1 + i); // 转换为0基索引
            if (row == null) {
                row = sheet.createRow(startRow - 1 + i); // 转换为0基索引
            }
            
            Map<String, Object> rowData = data.get(i);
            
            // 遍历表头列映射，填充数据
            for (Map.Entry<String, Integer> entry : headerColumnMap.entrySet()) {
                String templateField = entry.getKey(); // 模板中的字段名，如"姓名"
                int columnIndex = entry.getValue();    // 对应的列索引
                
                // 根据字段映射找到数据字段名
                String dataField = templateField; // 默认使用模板字段名作为数据字段名
                if (fieldMapping != null && fieldMapping.containsKey(templateField)) {
                    dataField = fieldMapping.get(templateField);
                }
                
                Cell cell = row.getCell(columnIndex);
                if (cell == null) {
                    cell = row.createCell(columnIndex);
                }
                
                // 设置单元格值
                Object value = rowData.get(dataField);
                setCellValue(cell, value);
            }
        }
    }

    /**
     * 查找列索引
     */
    private int findColumnIndex(Sheet sheet, String fieldName) {
        Row headerRow = sheet.getRow(0); // 假设标题行在第0行
        if (headerRow == null) {
            return -1;
        }
        
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null) {
                String cellValue = getCellValueAsString(cell);
                if (fieldName.equals(cellValue)) {
                    return i;
                }
                // 也检查是否包含占位符格式 {fieldName}
                if (cellValue != null && cellValue.startsWith("{") && cellValue.endsWith("}")) {
                    String placeholder = cellValue.substring(1, cellValue.length() - 1);
                    if (fieldName.equals(placeholder)) {
                        return i;
                    }
                }
            }
        }
        
        // 额外检查：如果在第1行（索引1）有占位符，也尝试查找
        Row placeholderRow = sheet.getRow(1);
        if (placeholderRow != null) {
            for (int i = 0; i < placeholderRow.getLastCellNum(); i++) {
                Cell cell = placeholderRow.getCell(i);
                if (cell != null) {
                    String cellValue = getCellValueAsString(cell);
                    if (cellValue != null && cellValue.startsWith("{") && cellValue.endsWith("}")) {
                        String placeholder = cellValue.substring(1, cellValue.length() - 1);
                        if (fieldName.equals(placeholder)) {
                            return i;
                        }
                    }
                }
            }
        }
        
        return -1;
    }

    /**
     * 获取单元格字符串值
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((long) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return cell.toString();
        }
    }

    /**
     * 设置单元格值
     */
    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setCellValue("");
            return;
        }
        
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * 检查模板中是否包含文档级占位符，并且数据中包含对应的字段
     */
    private boolean hasDocumentLevelPlaceholders(File templateFile, Map<String, Object> data) {
        try (FileInputStream fis = new FileInputStream(templateFile);
             org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis)) {
            
            // 遍历所有Sheet
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetIndex);
                
                // 遍历所有行和列
                for (org.apache.poi.ss.usermodel.Row row : sheet) {
                    if (row != null) {
                        for (org.apache.poi.ss.usermodel.Cell cell : row) {
                            if (cell != null && cell.getCellType() == CellType.STRING) {
                                String cellValue = cell.getStringCellValue();
                                if (cellValue != null) {
                                    // 检查是否包含${fieldName}格式的占位符
                                    for (String fieldName : data.keySet()) {
                                        String placeholder = "${" + fieldName + "}";
                                        if (cellValue.contains(placeholder)) {
                                            return true;
                                        }
                                    }
                                    // 检查是否包含{.fieldName}格式的占位符
                                    for (String fieldName : data.keySet()) {
                                        String placeholder = "{." + fieldName + "}";
                                        if (cellValue.contains(placeholder)) {
                                            return true;
                                        }
                                    }
                                    // 检查是否包含{fieldName}格式的占位符
                                    for (String fieldName : data.keySet()) {
                                        String placeholder = "{" + fieldName + "}";
                                        if (cellValue.contains(placeholder)) {
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error("检查文档级占位符失败：{}", templateFile.getAbsolutePath(), e);
        }
        
        return false;
    }

    /**
     * 更新Sheet名称
     */
    private void updateSheetName(File outputFile, String newSheetName) throws IOException {
        if (newSheetName == null || newSheetName.trim().isEmpty()) {
            return;
        }

        try (FileInputStream fis = new FileInputStream(outputFile);
             Workbook workbook = WorkbookFactory.create(fis)) {
            
            workbook.setSheetName(0, newSheetName);
            
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * 向Excel文件填充数据
     */
    private void fillDataToExcelFile(File outputFile, SingleSheetExportRequest request) throws IOException {
        // 读取已填充占位符的Excel文件
        try (FileInputStream fis = new FileInputStream(outputFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet sheet = workbook.getSheetAt(0);
            
            // 如果有数据需要填充，使用POI直接填充数据
            if (request.getData() != null && !request.getData().isEmpty()) {
                // 使用POI直接填充数据到指定行
                fillDataToSheet(sheet, request.getData(), request.getStartRow(), request.getFieldMapping());
            }
            
            // 如果需要更新Sheet名称
            if (request.getSheetName() != null && !request.getSheetName().isEmpty()) {
                workbook.setSheetName(0, request.getSheetName());
            }
            
            // 保存修改
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * 获取模板路径
     */
    private String getTemplatePath(String templateName) {
        if (templateName.startsWith("classpath:")) {
            try {
                ClassPathResource resource = new ClassPathResource(templateName.substring(10));
                return resource.getFile().getAbsolutePath();
            } catch (IOException e) {
                throw new RuntimeException("无法找到模板文件：" + templateName, e);
            }
        } else if (new File(templateName).exists()) {
            return templateName;
        } else {
            String path = excelConfig.getTemplatePath() + "/" + templateName + ".xlsx";
            return path.replace("//", "/");
        }
    }

    /**
     * 生成输出路径
     */
    private String generateOutputPath(String templateName, SingleSheetExportRequest request) {
        // 如果request不为null且用户指定了导出目录和文件名，则使用用户指定的
        if (request != null && 
            request.getExportDir() != null && !request.getExportDir().isEmpty() 
            && request.getExportFileName() != null && !request.getExportFileName().isEmpty()) {
            // 确保导出目录以分隔符结尾
            String exportDir = request.getExportDir().endsWith(File.separator) ? 
                request.getExportDir() : request.getExportDir() + File.separator;
            return exportDir + request.getExportFileName();
        } else {
            // 否则使用自动生成的文件名
            String fileName = fileNameGenerator.generateFileName(templateName, null);
            return excelConfig.getOutputPath() + "/" + fileName;
        }
    }

    /**
     * 向已填充占位符的文件添加多Sheet数据
     */
    private void addMultiSheetData(File outputFile, MultiSheetExportRequest request) throws Exception {
        // 读取已填充占位符的文件
        try (FileInputStream fis = new FileInputStream(outputFile);
             XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis)) {

            // 创建流式工作簿
            try (SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 1)) {
                
                // 删除原有的Sheet
                int numberOfSheets = sxssfWorkbook.getNumberOfSheets();
                for (int i = numberOfSheets - 1; i >= 0; i--) {
                    sxssfWorkbook.removeSheetAt(i);
                }

                // 遍历每个Sheet配置
                for (int i = 0; i < request.getSheetDataList().size(); i++) {
                    SheetDataConfig sheetConfig = request.getSheetDataList().get(i);
                    
                    // 创建或获取Sheet
                    String sheetName = sheetConfig.getSheetName() != null ? sheetConfig.getSheetName() : "Sheet" + (i + 1);
                    Sheet sheet = sxssfWorkbook.createSheet(sheetName);

                    // 填充数据到Sheet
                    fillDataToSheet(sheet, sheetConfig.getData(), sheetConfig.getStartRow(), 
                        sheetConfig.getFieldMapping());
                }

                // 写入文件
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    sxssfWorkbook.write(fos);
                }
                
                // 清理临时文件
                sxssfWorkbook.dispose();
            }
        }
    }
}