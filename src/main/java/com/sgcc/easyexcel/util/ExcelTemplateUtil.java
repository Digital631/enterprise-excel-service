    package com.sgcc.easyexcel.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.sgcc.easyexcel.config.ExcelConfig;
import com.sgcc.easyexcel.enums.ExcelErrorCode;
import com.sgcc.easyexcel.exception.ExcelExportException;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Excel模板工具类（线程安全）
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Component
public class ExcelTemplateUtil {

    private final ExcelConfig excelConfig;
    
    /**
     * 模板缓存（线程安全）
     */
    private final Map<String, List<String>> templateHeaderCache = new ConcurrentHashMap<>();

    public ExcelTemplateUtil(ExcelConfig excelConfig) {
        this.excelConfig = excelConfig;
    }

    /**
     * 解析模板路径（支持完整路径和模板名称）
     *
     * @param templatePath 模板标识
     * @return 模板完整路径
     */
    public String resolveTemplatePath(String templatePath) {
        if (templatePath == null || templatePath.trim().isEmpty()) {
            throw new ExcelExportException(ExcelErrorCode.TEMPLATE_PATH_EMPTY);
        }

        // 判断是否为完整路径（包含 / 或 \）
        if (templatePath.contains("/") || templatePath.contains("\\")) {
            return validateTemplatePath(templatePath);
        }

        // 模板名称，从配置目录查找
        String defaultDir = excelConfig.getTemplate().getDefaultDir();
        
        // 首先在模板目录中查找已存在的同名文件（不同扩展名）
        String existingFilePath = findExistingTemplateFile(templatePath);
        if (existingFilePath != null) {
            return existingFilePath;
        }
        
        String fullPath = defaultDir + File.separator + templatePath;

        // 如果没有扩展名，尝试添加
        if (!hasExtension(fullPath)) {
            fullPath = tryAddExtension(fullPath);
        }

        return validateTemplatePath(fullPath);
    }

    /**
     * 验证模板路径
     */
    private String validateTemplatePath(String templatePath) {
        File file = new File(templatePath);

        // 检查文件是否存在
        if (!file.exists()) {
            // 如果文件不存在，尝试其他可能的扩展名
            String actualPath = findTemplateWithAlternativeExtension(templatePath);
            if (actualPath != null) {
                file = new File(actualPath);
                templatePath = actualPath;
            } else {
                // 再次尝试在模板目录中查找已存在的同名文件
                String existingFilePath = findExistingTemplateFile(new File(templatePath).getName());
                if (existingFilePath != null) {
                    file = new File(existingFilePath);
                    templatePath = existingFilePath;
                } else {
                    throw new ExcelExportException(
                            ExcelErrorCode.TEMPLATE_NOT_FOUND,
                            "模板文件不存在：" + templatePath
                    );
                }
            }
        }

        // 检查是否为文件
        if (!file.isFile()) {
            throw new ExcelExportException(
                    ExcelErrorCode.TEMPLATE_FORMAT_ERROR,
                    "模板路径不是文件：" + templatePath
            );
        }

        // 检查文件扩展名
        List<String> allowedExtensions = excelConfig.getTemplate().getAllowedExtensions();
        String finalTemplatePath = templatePath;
        boolean validExtension = allowedExtensions.stream()
                .anyMatch(ext -> finalTemplatePath.toLowerCase().endsWith(ext));

        if (!validExtension) {
            throw new ExcelExportException(ExcelErrorCode.TEMPLATE_FORMAT_ERROR);
        }

        // 检查文件可读性
        if (!file.canRead()) {
            throw new ExcelExportException(
                    ExcelErrorCode.TEMPLATE_NO_PERMISSION,
                    "无权限读取模板文件：" + templatePath
            );
        }

        return templatePath;
    }

    /**
     * 查找具有不同扩展名的模板文件
     */
    private String findTemplateWithAlternativeExtension(String templatePath) {
        File originalFile = new File(templatePath);
        
        // 检查是否有扩展名
        int lastDotIndex = templatePath.lastIndexOf('.');
        if (lastDotIndex == -1) {
            // 没有扩展名，直接返回null，因为这种情况应该由tryAddExtension处理
            return null;
        }
        
        String basePath = templatePath.substring(0, lastDotIndex);
        List<String> allowedExtensions = excelConfig.getTemplate().getAllowedExtensions();
        
        // 首先尝试在模板目录中查找已存在的同名文件（不同扩展名）
        String defaultDir = excelConfig.getTemplate().getDefaultDir();
        File dir = new File(defaultDir);
        
        if (dir.exists() && dir.isDirectory()) {
            String originalFileName = new File(basePath).getName();
            String fileNameWithoutExt;
            if (originalFileName.lastIndexOf('.') != -1) {
                fileNameWithoutExt = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
            } else {
                fileNameWithoutExt = originalFileName;
            }
            
            // 查找目录中所有匹配的文件
            File[] files = dir.listFiles((d, name) -> 
                name.toLowerCase().startsWith(fileNameWithoutExt + ".") && 
                allowedExtensions.stream().anyMatch(ext -> name.toLowerCase().endsWith(ext))
            );
            
            if (files != null && files.length > 0) {
                // 如果找到匹配的文件，返回第一个匹配项
                return new File(defaultDir, files[0].getName()).getAbsolutePath();
            }
        }
        
        // 尝试所有允许的扩展名
        for (String ext : allowedExtensions) {
            String alternativePath = basePath + ext;
            File alternativeFile = new File(alternativePath);
            if (alternativeFile.exists() && alternativeFile.isFile()) {
                log.debug("找到模板文件：{} -> {}", templatePath, alternativePath);
                return alternativePath;
            }
        }
        
        return null; // 没有找到匹配的文件
    }

    /**
     * 读取模板表头（带缓存）
     *
     * @param templatePath 模板路径
     * @param sheetIndex Sheet索引
     * @return 表头字段列表
     */
    public List<String> readTemplateHeaders(String templatePath, int sheetIndex) {
        // 构建缓存key
        String cacheKey = templatePath + "#" + sheetIndex;

        // 从缓存获取
        if (excelConfig.getTemplate().getCache().getEnabled()) {
            List<String> cachedHeaders = templateHeaderCache.get(cacheKey);
            if (cachedHeaders != null) {
                log.debug("从缓存读取模板表头：{}", cacheKey);
                return new ArrayList<>(cachedHeaders);
            }
        }

        // 读取表头
        List<String> headers = doReadTemplateHeaders(templatePath, sheetIndex);

        // 存入缓存
        if (excelConfig.getTemplate().getCache().getEnabled()) {
            templateHeaderCache.put(cacheKey, headers);
            log.debug("模板表头已缓存：{}", cacheKey);
        }

        return headers;
    }

    /**
     * 实际读取表头
     */
    private List<String> doReadTemplateHeaders(String templatePath, int sheetIndex) {
        Integer headerRow = excelConfig.getFill().getHeaderRow();
        List<String> headers = new ArrayList<>();

        try {
            // 使用监听器读取表头行，注意：headRowNumber(0) 表示不跳过表头
            EasyExcel.read(templatePath, new AnalysisEventListener<Map<Integer, String>>() {
                @Override
                public void invoke(Map<Integer, String> data, AnalysisContext context) {
                    // rowIndex 从0开始，第1行的rowIndex为0
                    int currentRow = context.readRowHolder().getRowIndex();
                    log.debug("读取到第{}row, data={}", currentRow, data);
                    
                    // 读取指定的表头行（headerRow配置为1，对应rowIndex为0）
                    if (currentRow == headerRow - 1) {
                        data.values().stream()
                                .filter(Objects::nonNull)
                                .filter(h -> !h.trim().isEmpty())
                                .forEach(headers::add);
                        log.debug("表头读取完成：{}", headers);
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    // 读取完成
                }
            }).sheet(sheetIndex).headRowNumber(0).doRead();  // headRowNumber(0) 表示不跳过任何行

            if (headers.isEmpty()) {
                throw new ExcelExportException(
                        ExcelErrorCode.TEMPLATE_HEADER_ERROR,
                        "模板表头为空，请检查模板文件，确保第" + headerRow + "行有数据"
                );
            }

            log.info("读取模板表头成功：{}, headers={}", templatePath, headers);
            return headers;

        } catch (Exception e) {
            log.error("读取模板表头失败：{}", templatePath, e);
            throw new ExcelExportException(
                    ExcelErrorCode.TEMPLATE_HEADER_ERROR,
                    "读取模板表头失败：" + e.getMessage(),
                    e
            );
        }
    }

    /**
     * 建立字段映射关系
     *
     * @param headers 模板表头
     * @param customMapping 自定义映射
     * @param matchStrategy 匹配策略
     * @return 映射关系（表头 -> 字段名）
     */
    public Map<String, String> buildFieldMapping(
            List<String> headers,
            Map<String, String> customMapping,
            String matchStrategy
    ) {
        Map<String, String> mapping = new HashMap<>();

        for (String header : headers) {
            String fieldName;

            // 优先使用自定义映射
            if (customMapping != null && customMapping.containsKey(header)) {
                fieldName = customMapping.get(header);
            } else {
                // 根据策略匹配
                fieldName = matchFieldName(header, matchStrategy);
            }

            mapping.put(header, fieldName);
        }

        log.debug("字段映射：{}", mapping);
        return mapping;
    }

    /**
     * 匹配字段名
     */
    private String matchFieldName(String header, String strategy) {
        return switch (strategy) {
            case "exact" -> header; // 精确匹配
            case "ignore-case" -> header; // 忽略大小写（默认与精确一致）
            case "smart" -> smartMatch(header); // 智能匹配（驼峰转换等）
            default -> header;
        };
    }

    /**
     * 智能匹配（驼峰转换）
     */
    private String smartMatch(String header) {
        // 示例：将 "用户名称" 转为 "userName"
        // 这里简化处理，实际可根据业务需求扩展
        return header;
    }

    /**
     * 检查是否有扩展名
     */
    private boolean hasExtension(String path) {
        return path.contains(".") && path.lastIndexOf('.') > path.lastIndexOf(File.separator);
    }

    /**
     * 尝试添加扩展名
     */
    private String tryAddExtension(String path) {
        // 首先尝试在模板目录中查找已存在的同名文件（不同扩展名）
        String existingFilePath = findExistingTemplateFile(new File(path).getName());
        if (existingFilePath != null) {
            return existingFilePath;
        }
        
        // 按优先级尝试添加扩展名
        for (String ext : List.of(".xlsx", ".xls")) {
            String fullPath = path + ext;
            if (new File(fullPath).exists()) {
                return fullPath;
            }
        }
        
        // 默认返回 .xlsx
        return path + ".xlsx";
    }

    /**
     * 创建导出目录
     */
    public void ensureExportDir(String exportDir) {
        try {
            Path path = Paths.get(exportDir);
            if (!Files.exists(path)) {
                if (excelConfig.getExport().getAutoCreateDir()) {
                    Files.createDirectories(path);
                    log.info("自动创建导出目录：{}", exportDir);
                } else {
                    throw new ExcelExportException(
                            ExcelErrorCode.EXPORT_DIR_NOT_EXIST,
                            "导出目录不存在：" + exportDir
                    );
                }
            }

            // 检查目录权限
            if (!Files.isWritable(path)) {
                throw new ExcelExportException(
                        ExcelErrorCode.EXPORT_DIR_NO_PERMISSION,
                        "无权限写入导出目录：" + exportDir
                );
            }

        } catch (Exception e) {
            if (e instanceof ExcelExportException) {
                throw (ExcelExportException) e;
            }
            throw new ExcelExportException(
                    ExcelErrorCode.EXPORT_DIR_NOT_EXIST,
                    "创建导出目录失败：" + e.getMessage(),
                    e
            );
        }
    }

    /**
     * 清除模板缓存
     */
    public void clearCache() {
        templateHeaderCache.clear();
        log.info("模板缓存已清除");
    }

    /**
     * 填充模板（支持占位符和数据填充）
     *
     * @param templatePath 模板路径
     * @param outputPath 输出路径
     * @param data 数据列表
     * @param placeholders 占位符数据（用于填充文档级别的占位符）
     */
    public void fillTemplate(String templatePath, String outputPath, List<Map<String, Object>> data, Map<String, Object> placeholders) throws ExcelExportException {
        try {
            // 确保输出目录存在
            File outputFile = new File(outputPath);
            ensureExportDir(outputFile.getParent());

            // 如果模板中包含 {fieldName} 格式的占位符，需要预处理模板
            String processedTemplatePath = preprocessTemplate(templatePath, data);

            // 尝试使用EasyExcel处理，如果失败则使用POI直接处理
            boolean success = false;
            try {
                if (placeholders != null && !placeholders.isEmpty()) {
                    // 先填充占位符（文档级别的数据）
                    // 根据输出文件扩展名选择适当的写入方式
                    if (outputPath.toLowerCase().endsWith(".xlsx")) {
                        EasyExcel.write(outputPath, Map.class)
                                .withTemplate(processedTemplatePath)
                                .sheet()
                                .doFill(placeholders);
                    } else if (outputPath.toLowerCase().endsWith(".xls")) {
                        EasyExcel.write(outputPath, Map.class)
                                .withTemplate(processedTemplatePath)
                                .sheet()
                                .doFill(placeholders);
                    } else {
                        // 默认使用.xlsx格式
                        EasyExcel.write(outputPath, Map.class)
                                .withTemplate(processedTemplatePath)
                                .sheet()
                                .doFill(placeholders);
                    }
                    
                    // 然后基于已填充占位符的文件再次填充数据
                    if (data != null && !data.isEmpty()) {
                        // 根据输出文件扩展名选择适当的写入方式
                        if (outputPath.toLowerCase().endsWith(".xlsx")) {
                            EasyExcel.write(outputPath, Map.class)
                                    .withTemplate(outputPath) // 使用已填充占位符的文件作为模板
                                    .sheet()
                                    .doFill(data);
                        } else if (outputPath.toLowerCase().endsWith(".xls")) {
                            EasyExcel.write(outputPath, Map.class)
                                    .withTemplate(outputPath) // 使用已填充占位符的文件作为模板
                                    .sheet()
                                    .doFill(data);
                        } else {
                            // 默认使用.xlsx格式
                            EasyExcel.write(outputPath, Map.class)
                                    .withTemplate(outputPath) // 使用已填充占位符的文件作为模板
                                    .sheet()
                                    .doFill(data);
                        }
                    }
                } else {
                    // 直接填充数据
                    if (data != null && !data.isEmpty()) {
                        // 如果只有一行数据，我们直接用这一行数据替换模板中的占位符
                        if (data.size() == 1) {
                            // 使用POI直接处理单行数据的占位符替换
                            replacePlaceholdersInTemplate(processedTemplatePath, outputPath, data.get(0));
                        } else {
                            // 多行数据使用EasyExcel的doFill方法
                            // 根据输出文件扩展名选择适当的写入方式
                            if (outputPath.toLowerCase().endsWith(".xlsx")) {
                                EasyExcel.write(outputPath, Map.class)
                                        .withTemplate(processedTemplatePath)
                                        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                                        .sheet()
                                        .doFill(data);
                            } else if (outputPath.toLowerCase().endsWith(".xls")) {
                                EasyExcel.write(outputPath, Map.class)
                                        .withTemplate(processedTemplatePath)
                                        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                                        .sheet()
                                        .doFill(data);
                            } else {
                                // 默认使用.xlsx格式
                                EasyExcel.write(outputPath, Map.class)
                                        .withTemplate(processedTemplatePath)
                                        .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                                        .sheet()
                                        .doFill(data);
                            }
                        }
                    } else {
                        // 没有数据，直接复制模板
                        Path sourcePath = Paths.get(processedTemplatePath);
                        Path targetPath = Paths.get(outputPath);
                        Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
                    }
                }
                success = true;
            } catch (Exception easyExcelException) {
                log.warn("EasyExcel处理失败，尝试使用POI直接处理：{}", easyExcelException.getMessage());
            }

            // 如果EasyExcel处理失败，尝试使用POI直接处理
            if (!success) {
                log.info("使用POI直接处理Excel文件：{} -> {}", processedTemplatePath, outputPath);
                
                // 根据模板文件扩展名选择适当的处理方式
                if (processedTemplatePath.toLowerCase().endsWith(".xlsx")) {
                    try (FileInputStream fis = new FileInputStream(processedTemplatePath);
                         XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
                        
                        // 如果有占位符，先处理占位符
                        if (placeholders != null && !placeholders.isEmpty()) {
                            processPlaceholdersInWorkbook(workbook, placeholders);
                        }
                        
                        // 如果有数据，填充数据
                        if (data != null && !data.isEmpty()) {
                            processDataFillInWorkbook(workbook, data);
                        }
                        
                        // 保存到输出文件
                        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                            workbook.write(fos);
                        }
                    }
                } else if (processedTemplatePath.toLowerCase().endsWith(".xls")) {
                    try (FileInputStream fis = new FileInputStream(processedTemplatePath);
                         HSSFWorkbook workbook = new HSSFWorkbook(fis)) {
                        
                        // 如果有占位符，先处理占位符
                        if (placeholders != null && !placeholders.isEmpty()) {
                            processPlaceholdersInWorkbook(workbook, placeholders);
                        }
                        
                        // 如果有数据，填充数据
                        if (data != null && !data.isEmpty()) {
                            processDataFillInWorkbook(workbook, data);
                        }
                        
                        // 保存到输出文件
                        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                            workbook.write(fos);
                        }
                    }
                } else {
                    throw new ExcelExportException(ExcelErrorCode.TEMPLATE_FORMAT_ERROR, 
                        "不支持的模板格式，仅支持.xlsx和.xls文件");
                }
            }
            
            log.info("模板填充完成：{} -> {}", templatePath, outputPath);
        } catch (Exception e) {
            log.error("模板填充失败：{}", templatePath, e);
            throw new ExcelExportException(ExcelErrorCode.FILE_WRITE_ERROR, 
                "模板填充失败：" + e.getMessage());
        }
    }
        
    /**
     * 预处理模板，将 {fieldName} 格式转换为 ${fieldName} 格式
     */
    private String preprocessTemplate(String templatePath, List<Map<String, Object>> data) throws ExcelExportException {
        if (data == null || data.isEmpty()) {
            return templatePath; // 如果没有数据，无需预处理
        }
            
        // 创建临时文件用于存储预处理后的模板
        String tempDir = System.getProperty("java.io.tmpdir");
        String fileName = "preprocessed_" + System.currentTimeMillis() + "_" + new File(templatePath).getName();
        String tempTemplatePath = tempDir + File.separator + fileName;
            
        try {
            // 复制原始模板到临时文件
            Path sourcePath = Paths.get(templatePath);
            Path targetPath = Paths.get(tempTemplatePath);
            Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
                
            // 检查数据中是否包含 {fieldName} 格式的占位符
            boolean hasCurlyBracePlaceholders = false;
            for (Map<String, Object> row : data) {
                for (Object value : row.values()) {
                    if (value instanceof String && ((String) value).startsWith("{") && ((String) value).endsWith("}")) {
                        hasCurlyBracePlaceholders = true;
                        break;
                    }
                }
                if (hasCurlyBracePlaceholders) {
                    break;
                }
            }
                
            if (hasCurlyBracePlaceholders) {
                // 根据文件扩展名选择适当的处理方式
                if (tempTemplatePath.toLowerCase().endsWith(".xlsx")) {
                    try (FileInputStream fis = new FileInputStream(tempTemplatePath);
                         XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
                            
                        // 遍历所有Sheet
                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            Sheet sheet = workbook.getSheetAt(i);
                                
                            // 遍历所有行和列，查找 {fieldName} 格式的占位符并替换为 ${fieldName}
                            for (Row row : sheet) {
                                if (row != null) {
                                    for (Cell cell : row) {
                                        if (cell != null && cell.getCellType() == CellType.STRING) {
                                            String cellValue = cell.getStringCellValue();
                                            if (cellValue != null) {
                                                // 查找 {fieldName} 格式的占位符并替换为 ${fieldName}
                                                for (Map<String, Object> dataRow : data) {
                                                    for (String key : dataRow.keySet()) {
                                                        String oldPlaceholder = "{" + key + "}";
                                                        String newPlaceholder = "${" + key + "}";
                                                        if (cellValue.contains(oldPlaceholder)) {
                                                            cellValue = cellValue.replace(oldPlaceholder, newPlaceholder);
                                                            cell.setCellValue(cellValue);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                            
                        // 保存到临时文件
                        try (FileOutputStream fos = new FileOutputStream(tempTemplatePath)) {
                            workbook.write(fos);
                        }
                    }
                } else if (tempTemplatePath.toLowerCase().endsWith(".xls")) {
                    try (FileInputStream fis = new FileInputStream(tempTemplatePath);
                         HSSFWorkbook workbook = new HSSFWorkbook(fis)) {
                            
                        // 遍历所有Sheet
                        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                            Sheet sheet = workbook.getSheetAt(i);
                                
                            // 遍历所有行和列，查找 {fieldName} 格式的占位符并替换为 ${fieldName}
                            for (Row row : sheet) {
                                if (row != null) {
                                    for (Cell cell : row) {
                                        if (cell != null && cell.getCellType() == CellType.STRING) {
                                            String cellValue = cell.getStringCellValue();
                                            if (cellValue != null) {
                                                // 查找 {fieldName} 格式的占位符并替换为 ${fieldName}
                                                for (Map<String, Object> dataRow : data) {
                                                    for (String key : dataRow.keySet()) {
                                                        String oldPlaceholder = "{" + key + "}";
                                                        String newPlaceholder = "${" + key + "}";
                                                        if (cellValue.contains(oldPlaceholder)) {
                                                            cellValue = cellValue.replace(oldPlaceholder, newPlaceholder);
                                                            cell.setCellValue(cellValue);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                            
                        // 保存到临时文件
                        try (FileOutputStream fos = new FileOutputStream(tempTemplatePath)) {
                            workbook.write(fos);
                        }
                    }
                } else {
                    throw new ExcelExportException(ExcelErrorCode.TEMPLATE_FORMAT_ERROR, 
                        "不支持的模板格式，仅支持.xlsx和.xls文件");
                }
            }
                
            return tempTemplatePath;
        } catch (Exception e) {
            log.error("模板预处理失败：{}", templatePath, e);
            throw new ExcelExportException(ExcelErrorCode.TEMPLATE_PROCESS_ERROR, 
                "模板预处理失败：" + e.getMessage());
        }
    }
        
    /**
     * 使用POI直接替换模板中的占位符
     */
    private void replacePlaceholdersInTemplate(String templatePath, String outputPath, Map<String, Object> placeholders) throws ExcelExportException {
        try {
            // 根据文件扩展名选择适当的处理方式
            if (templatePath.toLowerCase().endsWith(".xlsx")) {
                try (FileInputStream fis = new FileInputStream(templatePath);
                     XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
                        
                    // 处理占位符
                    processPlaceholdersInWorkbook(workbook, placeholders);
                        
                    // 保存到输出文件
                    try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                        workbook.write(fos);
                    }
                }
            } else if (templatePath.toLowerCase().endsWith(".xls")) {
                try (FileInputStream fis = new FileInputStream(templatePath);
                     HSSFWorkbook workbook = new HSSFWorkbook(fis)) {
                        
                    // 处理占位符
                    processPlaceholdersInWorkbook(workbook, placeholders);
                        
                    // 保存到输出文件
                    try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                        workbook.write(fos);
                    }
                }
            } else {
                throw new ExcelExportException(ExcelErrorCode.TEMPLATE_FORMAT_ERROR, 
                    "不支持的模板格式，仅支持.xlsx和.xls文件");
            }
        } catch (Exception e) {
            log.error("占位符替换失败：{} -> {}", templatePath, outputPath, e);
            throw new ExcelExportException(ExcelErrorCode.FILE_WRITE_ERROR, 
                "占位符替换失败：" + e.getMessage());
        }
    }
        
    /**
     * 在工作簿中处理占位符
     */
    private void processPlaceholdersInWorkbook(Workbook workbook, Map<String, Object> placeholders) {
        // 遍历所有Sheet
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            
            // 遍历所有行和列
            for (Row row : sheet) {
                if (row != null) {
                    for (Cell cell : row) {
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            if (cellValue != null) {
                                // 替换所有占位符
                                for (Map.Entry<String, Object> entry : placeholders.entrySet()) {
                                    String placeholder = "${" + entry.getKey() + "}";
                                    String value = entry.getValue() != null ? entry.getValue().toString() : "";
                                    if (cellValue.contains(placeholder)) {
                                        cellValue = cellValue.replace(placeholder, value);
                                        cell.setCellValue(cellValue);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    /**
     * 在工作簿中填充数据
     */
    private void processDataFillInWorkbook(Workbook workbook, List<Map<String, Object>> data) {
        // 这里实现数据填充逻辑，基于第一个Sheet进行数据填充
        Sheet sheet = workbook.getSheetAt(0);
        
        // 找到数据起始行（通常在模板中是空行或特定标记行）
        int startRow = findDataStartRow(sheet);
        
        // 填充数据
        for (int i = 0; i < data.size(); i++) {
            Map<String, Object> rowData = data.get(i);
            Row row = sheet.getRow(startRow + i);
            if (row == null) {
                row = sheet.createRow(startRow + i);
            }
            
            int cellIndex = 0;
            for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                Cell cell = row.getCell(cellIndex);
                if (cell == null) {
                    cell = row.createCell(cellIndex);
                }
                
                Object value = entry.getValue();
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Number) {
                    if (value instanceof Integer || value instanceof Long) {
                        cell.setCellValue(((Number) value).longValue());
                    } else {
                        cell.setCellValue(((Number) value).doubleValue());
                    }
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                } else {
                    cell.setCellValue(value != null ? value.toString() : "");
                }
                cellIndex++;
            }
        }
    }
    
    /**
     * 查找数据起始行
     */
    private int findDataStartRow(Sheet sheet) {
        // 简单实现：查找第一个空行或特定标记行，这里默认从第2行开始
        return 1; // 从第2行开始（0-based）
    }
    
    /**
     * 在模板目录中查找已存在的同名文件（不同扩展名）
     */
    private String findExistingTemplateFile(String templateName) {
        String defaultDir = excelConfig.getTemplate().getDefaultDir();
        File dir = new File(defaultDir);
        
        if (dir.exists() && dir.isDirectory()) {
            String fileNameWithoutExt;
            if (templateName.lastIndexOf('.') != -1) {
                fileNameWithoutExt = templateName.substring(0, templateName.lastIndexOf('.'));
            } else {
                fileNameWithoutExt = templateName;
            }
            
            log.debug("在模板目录中查找文件：{}，基础名称：{}", templateName, fileNameWithoutExt);
            
            // 查找目录中所有匹配的文件
            File[] files = dir.listFiles((d, name) -> {
                String lowerName = name.toLowerCase();
                String lowerBaseName = fileNameWithoutExt.toLowerCase();
                boolean startsWithBaseName = lowerName.startsWith(lowerBaseName + ".");
                boolean hasValidExtension = lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
                log.debug("检查文件：{}，是否匹配基础名称：{}，是否有有效扩展名：{}", name, startsWithBaseName, hasValidExtension);
                return startsWithBaseName && hasValidExtension;
            });
            
            if (files != null && files.length > 0) {
                log.debug("找到匹配的文件：{}", files[0].getName());
                // 如果找到匹配的文件，返回完整路径
                return new File(defaultDir, files[0].getName()).getAbsolutePath();
            }
        }
        
        return null; // 没有找到匹配的文件
    }
}
