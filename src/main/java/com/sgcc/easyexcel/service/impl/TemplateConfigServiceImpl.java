package com.sgcc.easyexcel.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.sgcc.easyexcel.config.ExcelConfig;
import com.sgcc.easyexcel.dto.request.TemplateFieldMappingRequest;
import com.sgcc.easyexcel.dto.response.TemplateInfoResponse;
import com.sgcc.easyexcel.service.TemplateConfigService;
import com.sgcc.easyexcel.util.ExcelTemplateUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 模板配置服务实现
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Service
public class TemplateConfigServiceImpl implements TemplateConfigService {

    private final ExcelConfig excelConfig;
    private final ExcelTemplateUtil excelTemplateUtil;

    /**
     * 字段映射配置缓存（模板名称 -> 字段映射）
     */
    private final Map<String, Map<String, String>> fieldMappingCache = new ConcurrentHashMap<>();

    public TemplateConfigServiceImpl(ExcelConfig excelConfig, ExcelTemplateUtil excelTemplateUtil) {
        this.excelConfig = excelConfig;
        this.excelTemplateUtil = excelTemplateUtil;
    }

    @Override
    public List<TemplateInfoResponse> getAllTemplates() {
        List<TemplateInfoResponse> templates = new ArrayList<>();
        String templateDir = excelConfig.getTemplate().getDefaultDir();
        File dir = new File(templateDir);

        if (!dir.exists() || !dir.isDirectory()) {
            log.warn("模板目录不存在：{}" , templateDir);
            return templates;
        }

        File[] files = dir.listFiles((d, name) -> {
            String lowerName = name.toLowerCase();
            return lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
        });

        if (files != null) {
            for (File file : files) {
                try {
                    templates.add(buildTemplateInfo(file));
                } catch (Exception e) {
                    log.error("读取模板信息失败：{}" , file.getName(), e);
                }
            }
        }

        // 按最后修改时间倒序排列
        templates.sort((a, b) -> b.getLastModified().compareTo(a.getLastModified()));
        return templates;
    }

    @Override
    public TemplateInfoResponse getTemplateInfo(String templateName) {
        String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
        File file = new File(templatePath);
        
        if (!file.exists()) {
            throw new RuntimeException("模板文件不存在：" + templateName);
        }

        return buildTemplateInfo(file);
    }

    @Override
    public TemplateInfoResponse uploadTemplate(MultipartFile file) throws IOException {
        String templateDir = excelConfig.getTemplate().getDefaultDir();
        
        // 确保目录存在
        Path dirPath = Paths.get(templateDir);
        if (!Files.exists(dirPath)) {
            Files.createDirectories(dirPath);
        }

        // 检查文件扩展名
        String originalFilename = file.getOriginalFilename();
        if (originalFilename == null || 
            (!originalFilename.toLowerCase().endsWith(".xlsx") && 
             !originalFilename.toLowerCase().endsWith(".xls"))) {
            throw new RuntimeException("只支持 .xlsx 或 .xls 格式的Excel文件");
        }

        // 保存文件
        String fileName = originalFilename;
        Path targetPath = Paths.get(templateDir, fileName);
        
        // 如果文件已存在，先删除
        if (Files.exists(targetPath)) {
            Files.delete(targetPath);
        }
        
        Files.copy(file.getInputStream(), targetPath, StandardCopyOption.REPLACE_EXISTING);
        log.info("模板文件上传成功：{}" , fileName);

        return buildTemplateInfo(targetPath.toFile());
    }

    @Override
    public boolean deleteTemplate(String templateName) {
        try {
            String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
            File file = new File(templatePath);
            
            if (file.exists() && file.delete()) {
                // 删除对应的字段映射配置
                fieldMappingCache.remove(templateName);
                log.info("模板文件删除成功：{}" , templateName);
                return true;
            }
            return false;
        } catch (Exception e) {
            log.error("删除模板文件失败：{}" , templateName, e);
            return false;
        }
    }

    @Override
    public boolean saveFieldMapping(TemplateFieldMappingRequest request) {
        String templateName = request.getTemplateName();
        
        // 验证模板是否存在
        try {
            excelTemplateUtil.resolveTemplatePath(templateName);
        } catch (Exception e) {
            throw new RuntimeException("模板不存在：" + templateName);
        }

        // 保存字段映射
        fieldMappingCache.put(templateName, request.getFieldMapping());
        log.info("字段映射配置保存成功：{}" , templateName);
        return true;
    }

    @Override
    public Map<String, String> getFieldMapping(String templateName) {
        return fieldMappingCache.getOrDefault(templateName, new HashMap<>());
    }

    @Override
    public List<Map<String, Object>> previewTemplateData(String templateName, int sheetIndex) {
        String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
        List<Map<String, Object>> previewData = new ArrayList<>();
        
        try {
            // 读取前5行数据作为预览
            EasyExcel.read(templatePath, new AnalysisEventListener<Map<Integer, String>>() {
                private int rowCount = 0;
                private List<String> headers = new ArrayList<>();

                @Override
                public void invoke(Map<Integer, String> data, AnalysisContext context) {
                    int currentRow = context.readRowHolder().getRowIndex();
                    
                    // 第一行是表头
                    if (currentRow == 0) {
                        data.values().forEach(headers::add);
                    } else if (rowCount < 5) { // 只读取前5行数据
                        Map<String, Object> rowData = new LinkedHashMap<>();
                        int colIndex = 0;
                        for (String header : headers) {
                            rowData.put(header, data.get(colIndex));
                            colIndex++;
                        }
                        previewData.add(rowData);
                        rowCount++;
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    // 读取完成
                }
            }).sheet(sheetIndex).headRowNumber(0).doRead();
            
        } catch (Exception e) {
            log.error("预览模板数据失败：{}" , templateName, e);
        }

        return previewData;
    }

    @Override
    public Map<String, Object> validateTemplate(String templateName) {
        Map<String, Object> result = new HashMap<>();
        List<String> errors = new ArrayList<>();
        List<String> warnings = new ArrayList<>();

        try {
            String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
            File file = new File(templatePath);

            // 检查文件是否存在且可读
            if (!file.exists()) {
                errors.add("模板文件不存在");
            } else if (!file.canRead()) {
                errors.add("模板文件无法读取");
            } else {
                // 检查文件是否能正常打开
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook workbook = WorkbookFactory.create(fis)) {
                    
                    int sheetCount = workbook.getNumberOfSheets();
                    result.put("sheetCount", sheetCount);
                    
                    if (sheetCount == 0) {
                        errors.add("模板文件不包含任何Sheet");
                    } else {
                        // 检查每个Sheet
                        for (int i = 0; i < sheetCount; i++) {
                            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(i);
                            int rowCount = sheet.getPhysicalNumberOfRows();
                            
                            if (rowCount == 0) {
                                warnings.add("Sheet " + (i + 1) + " (" + sheet.getSheetName() + ") 为空");
                            } else if (rowCount == 1) {
                                warnings.add("Sheet " + (i + 1) + " (" + sheet.getSheetName() + ") 只有表头行，没有数据行");
                            }
                        }
                    }
                }
            }

            // 检查是否有字段映射配置
            Map<String, String> fieldMapping = fieldMappingCache.get(templateName);
            if (fieldMapping == null || fieldMapping.isEmpty()) {
                warnings.add("尚未配置字段映射，导出时可能需要手动指定");
            } else {
                result.put("fieldMappingCount", fieldMapping.size());
            }

        } catch (Exception e) {
            errors.add("验证失败：" + e.getMessage());
        }

        result.put("valid", errors.isEmpty());
        result.put("errors", errors);
        result.put("warnings", warnings);
        
        return result;
    }

    /**
     * 构建模板信息
     */
    private TemplateInfoResponse buildTemplateInfo(File file) {
        TemplateInfoResponse.TemplateInfoResponseBuilder builder = TemplateInfoResponse.builder();
        
        builder.fileName(file.getName());
        builder.filePath(file.getAbsolutePath());
        builder.fileSize(file.length());
        builder.fileSizeFormatted(formatFileSize(file.length()));
        builder.lastModified(LocalDateTime.ofInstant(
                Instant.ofEpochMilli(file.lastModified()), 
                ZoneId.systemDefault()));

        // 读取Sheet信息
        List<TemplateInfoResponse.SheetInfo> sheets = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(i);
                
                // 读取表头
                List<String> headers = new ArrayList<>();
                org.apache.poi.ss.usermodel.Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    for (org.apache.poi.ss.usermodel.Cell cell : headerRow) {
                        String value = "";
                        switch (cell.getCellType()) {
                            case STRING:
                                value = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                value = String.valueOf(cell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                value = String.valueOf(cell.getBooleanCellValue());
                                break;
                            default:
                                value = "";
                        }
                        headers.add(value.trim());
                    }
                }
                
                TemplateInfoResponse.SheetInfo sheetInfo = TemplateInfoResponse.SheetInfo.builder()
                        .sheetIndex(i)
                        .sheetName(sheet.getSheetName())
                        .dataRowCount(sheet.getPhysicalNumberOfRows())
                        .headers(headers)
                        .build();
                sheets.add(sheetInfo);
            }
        } catch (Exception e) {
            log.error("读取模板Sheet信息失败：{}" , file.getName(), e);
        }
        builder.sheets(sheets);

        // 检查是否有字段映射配置
        Map<String, String> fieldMapping = fieldMappingCache.get(file.getName());
        builder.hasFieldMapping(fieldMapping != null && !fieldMapping.isEmpty());
        builder.fieldMapping(fieldMapping);

        return builder.build();
    }

    /**
     * 格式化文件大小
     */
    private String formatFileSize(long size) {
        if (size < 1024) {
            return size + " B";
        } else if (size < 1024 * 1024) {
            return String.format("%.2f KB", size / 1024.0);
        } else {
            return String.format("%.2f MB", size / (1024.0 * 1024));
        }
    }
}
