package com.sgcc.easyexcel.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.sgcc.easyexcel.config.ExcelConfig;
import com.sgcc.easyexcel.dto.request.TemplateSaveRequest;
import com.sgcc.easyexcel.service.ExcelDesignerService;
import com.sgcc.easyexcel.util.ExcelTemplateUtil;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel模板设计器服务实现
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Service
public class ExcelDesignerServiceImpl implements ExcelDesignerService {

    private final ExcelConfig excelConfig;
    private final ExcelTemplateUtil excelTemplateUtil;

    // 占位符正则表达式: ${fieldName} 或 {fieldName}
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$?\\{([a-zA-Z_][a-zA-Z0-9_]*)\\}");

    public ExcelDesignerServiceImpl(ExcelConfig excelConfig, ExcelTemplateUtil excelTemplateUtil) {
        this.excelConfig = excelConfig;
        this.excelTemplateUtil = excelTemplateUtil;
    }

    @Override
    public List<Map<String, Object>> getTemplateList() {
        List<Map<String, Object>> templates = new ArrayList<>();
        String templateDir = excelConfig.getTemplate().getDefaultDir();
        File dir = new File(templateDir);

        if (!dir.exists() || !dir.isDirectory()) {
            log.warn("模板目录不存在：{}", templateDir);
            return templates;
        }

        File[] files = dir.listFiles((d, name) -> {
            String lowerName = name.toLowerCase();
            return lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
        });

        if (files != null) {
            for (File file : files) {
                Map<String, Object> template = new HashMap<>();
                template.put("fileName", file.getName());
                template.put("fileSize", file.length());
                template.put("fileSizeFormatted", formatFileSize(file.length()));
                template.put("lastModified", LocalDateTime.ofInstant(
                        Instant.ofEpochMilli(file.lastModified()),
                        ZoneId.systemDefault()));
                
                // 提取占位符
                try {
                    List<String> placeholders = extractPlaceholdersFromFile(file);
                    template.put("placeholders", placeholders);
                    template.put("placeholderCount", placeholders.size());
                } catch (Exception e) {
                    template.put("placeholders", new ArrayList<>());
                    template.put("placeholderCount", 0);
                }
                
                templates.add(template);
            }
        }

        // 按修改时间倒序
        templates.sort((a, b) -> ((LocalDateTime) b.get("lastModified"))
                .compareTo((LocalDateTime) a.get("lastModified")));
        
        return templates;
    }

    @Override
    public Map<String, Object> loadTemplate(String templateName) {
        try {
            String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
            File file = new File(templatePath);
            
            if (!file.exists()) {
                throw new RuntimeException("模板文件不存在：" + templateName);
            }

            // 读取Excel文件并转换为Luckysheet格式
            return convertExcelToLuckysheet(file);
        } catch (Exception e) {
            log.error("加载模板失败：{}", templateName, e);
            throw new RuntimeException("加载模板失败：" + e.getMessage(), e);
        }
    }

    @Override
    public String saveTemplate(TemplateSaveRequest request) {
        try {
            String templateDir = excelConfig.getTemplate().getDefaultDir();
            Path dirPath = Paths.get(templateDir);
            
            if (!Files.exists(dirPath)) {
                Files.createDirectories(dirPath);
            }

            // 确保文件名有正确的扩展名
            String fileName = request.getTemplateName();
            if (!fileName.toLowerCase().endsWith(".xlsx") && !fileName.toLowerCase().endsWith(".xls")) {
                fileName = fileName + ".xlsx";
            }

            String filePath = templateDir + File.separator + fileName;
            File file = new File(filePath);

            // 检查是否覆盖
            if (file.exists() && !Boolean.TRUE.equals(request.getOverwrite())) {
                throw new RuntimeException("模板已存在，请设置overwrite=true覆盖");
            }

            // 将Luckysheet数据转换为Excel并保存
            convertLuckysheetToExcel(request.getSheets(), filePath);

            log.info("模板保存成功：{}", filePath);
            return filePath;
        } catch (Exception e) {
            log.error("保存模板失败", e);
            throw new RuntimeException("保存模板失败：" + e.getMessage(), e);
        }
    }

    @Override
    public Map<String, Object> importExcel(MultipartFile file) throws IOException {
        // 保存到临时文件
        String tempDir = System.getProperty("java.io.tmpdir");
        String tempFileName = "import_" + System.currentTimeMillis() + "_" + file.getOriginalFilename();
        String tempFilePath = tempDir + File.separator + tempFileName;
        
        File tempFile = new File(tempFilePath);
        file.transferTo(tempFile);

        try {
            // 转换为Luckysheet格式
            return convertExcelToLuckysheet(tempFile);
        } finally {
            // 清理临时文件
            tempFile.delete();
        }
    }

    @Override
    public void exportExcel(String templateName, HttpServletResponse response) throws IOException {
        String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
        File file = new File(templatePath);
        
        if (!file.exists()) {
            response.setStatus(404);
            response.getWriter().write("模板文件不存在");
            return;
        }

        // 设置响应头
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=" + 
                new String(templateName.getBytes("UTF-8"), "ISO-8859-1"));
        response.setContentLength((int) file.length());

        // 写入响应流
        try (FileInputStream fis = new FileInputStream(file);
             OutputStream os = response.getOutputStream()) {
            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = fis.read(buffer)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
        }
    }

    @Override
    public boolean deleteTemplate(String templateName) {
        try {
            String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
            File file = new File(templatePath);
            
            if (file.exists()) {
                return file.delete();
            }
            return false;
        } catch (Exception e) {
            log.error("删除模板失败：{}", templateName, e);
            return false;
        }
    }

    @Override
    public List<String> extractPlaceholders(String templateName) {
        try {
            String templatePath = excelTemplateUtil.resolveTemplatePath(templateName);
            File file = new File(templatePath);
            return extractPlaceholdersFromFile(file);
        } catch (Exception e) {
            log.error("提取占位符失败：{}", templateName, e);
            return new ArrayList<>();
        }
    }

    @Override
    public Map<String, Object> previewTemplate(String templateName, Map<String, Object> sampleData) {
        // 这里可以实现模板预览功能
        // 使用示例数据填充模板并返回预览结果
        Map<String, Object> preview = new HashMap<>();
        preview.put("templateName", templateName);
        preview.put("sampleData", sampleData);
        preview.put("previewTime", LocalDateTime.now());
        return preview;
    }

    /**
     * 从文件中提取占位符
     */
    private List<String> extractPlaceholdersFromFile(File file) throws IOException {
        Set<String> placeholders = new HashSet<>();
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                
                for (Row row : sheet) {
                    if (row == null) continue;
                    
                    for (Cell cell : row) {
                        if (cell == null) continue;
                        
                        String value = getCellValueAsString(cell);
                        if (value != null && !value.isEmpty()) {
                            Matcher matcher = PLACEHOLDER_PATTERN.matcher(value);
                            while (matcher.find()) {
                                placeholders.add(matcher.group(1));
                            }
                        }
                    }
                }
            }
        }
        
        return new ArrayList<>(placeholders);
    }

    /**
     * 将Excel文件转换为Luckysheet格式
     */
    private Map<String, Object> convertExcelToLuckysheet(File file) throws IOException {
        Map<String, Object> result = new HashMap<>();
        List<Map<String, Object>> sheets = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {
            
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                Map<String, Object> sheetData = convertSheetToLuckysheet(sheet, i);
                sheets.add(sheetData);
            }
        }
        
        result.put("sheets", sheets);
        result.put("title", file.getName());
        return result;
    }

    /**
     * 将POI Sheet转换为Luckysheet格式
     */
    private Map<String, Object> convertSheetToLuckysheet(Sheet sheet, int index) {
        Map<String, Object> sheetData = new HashMap<>();
        sheetData.put("name", sheet.getSheetName());
        sheetData.put("index", String.valueOf(index));
        sheetData.put("status", index == 0 ? "1" : "0");
        
        // 转换单元格数据
        Map<String, Object> celldata = new HashMap<>();
        List<Map<String, Object>> cellDataList = new ArrayList<>();
        
        int maxRow = 0;
        int maxCol = 0;
        
        for (Row row : sheet) {
            if (row == null) continue;
            
            int rowNum = row.getRowNum();
            maxRow = Math.max(maxRow, rowNum);
            
            for (Cell cell : row) {
                if (cell == null) continue;
                
                int colNum = cell.getColumnIndex();
                maxCol = Math.max(maxCol, colNum);
                
                Map<String, Object> cellData = convertCellToLuckysheet(cell, rowNum, colNum);
                if (cellData != null) {
                    cellDataList.add(cellData);
                }
            }
        }
        
        sheetData.put("celldata", cellDataList);
        sheetData.put("row", maxRow + 1);
        sheetData.put("column", maxCol + 1);
        
        // 合并单元格信息
        List<Map<String, Object>> merges = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            Map<String, Object> merge = new HashMap<>();
            merge.put("r", region.getFirstRow());
            merge.put("c", region.getFirstColumn());
            merge.put("rs", region.getLastRow() - region.getFirstRow() + 1);
            merge.put("cs", region.getLastColumn() - region.getFirstColumn() + 1);
            merges.add(merge);
        }
        sheetData.put("config", Map.of("merge", merges));
        
        return sheetData;
    }

    /**
     * 将POI Cell转换为Luckysheet格式
     */
    private Map<String, Object> convertCellToLuckysheet(Cell cell, int r, int c) {
        Map<String, Object> cellData = new HashMap<>();
        cellData.put("r", r);
        cellData.put("c", c);
        
        Map<String, Object> v = new HashMap<>();
        
        switch (cell.getCellType()) {
            case STRING:
                v.put("v", cell.getStringCellValue());
                v.put("ct", Map.of("fa", "General", "t", "g"));
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    v.put("v", cell.getDateCellValue().toString());
                    v.put("ct", Map.of("fa", "yyyy-MM-dd", "t", "d"));
                } else {
                    v.put("v", cell.getNumericCellValue());
                    v.put("ct", Map.of("fa", "General", "t", "n"));
                }
                break;
            case BOOLEAN:
                v.put("v", cell.getBooleanCellValue());
                v.put("ct", Map.of("fa", "General", "t", "b"));
                break;
            case FORMULA:
                v.put("v", cell.getCellFormula());
                v.put("f", cell.getCellFormula());
                v.put("ct", Map.of("fa", "General", "t", "g"));
                break;
            default:
                v.put("v", "");
                v.put("ct", Map.of("fa", "General", "t", "g"));
        }
        
        // 样式信息
        CellStyle style = cell.getCellStyle();
        Map<String, Object> styleMap = new HashMap<>();
        
        // 字体
        Font font = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());
        if (font != null) {
            Map<String, Object> fontMap = new HashMap<>();
            fontMap.put("fs", font.getFontHeightInPoints());
            fontMap.put("bl", font.getBold() ? 1 : 0);
            styleMap.put("ff", fontMap);
        }
        
        // 对齐
        styleMap.put("ht", style.getAlignment().getCode());
        styleMap.put("vt", style.getVerticalAlignment().getCode());
        
        if (!styleMap.isEmpty()) {
            v.put("s", styleMap);
        }
        
        cellData.put("v", v);
        return cellData;
    }

    /**
     * 将Luckysheet数据转换为Excel文件
     */
    private void convertLuckysheetToExcel(List<Map<String, Object>> sheets, String filePath) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            
            for (Map<String, Object> sheetData : sheets) {
                String sheetName = (String) sheetData.getOrDefault("name", "Sheet");
                Sheet sheet = workbook.createSheet(sheetName);
                
                // 写入单元格数据
                List<Map<String, Object>> cellDataList = (List<Map<String, Object>>) sheetData.get("celldata");
                if (cellDataList != null) {
                    for (Map<String, Object> cellData : cellDataList) {
                        try {
                            int r = ((Number) cellData.get("r")).intValue();
                            int c = ((Number) cellData.get("c")).intValue();
                            Object vObj = cellData.get("v");
                            
                            // v可能是Map或直接的值
                            if (vObj != null) {
                                Row row = sheet.getRow(r);
                                if (row == null) {
                                    row = sheet.createRow(r);
                                }
                                
                                Cell cell = row.createCell(c);
                                
                                if (vObj instanceof Map) {
                                    Map<String, Object> v = (Map<String, Object>) vObj;
                                    setCellValueFromLuckysheet(cell, v);
                                } else {
                                    // 直接设置值
                                    if (vObj instanceof Number) {
                                        cell.setCellValue(((Number) vObj).doubleValue());
                                    } else if (vObj instanceof Boolean) {
                                        cell.setCellValue((Boolean) vObj);
                                    } else {
                                        cell.setCellValue(vObj.toString());
                                    }
                                }
                            }
                        } catch (Exception e) {
                            log.warn("处理单元格数据失败: {}", cellData, e);
                        }
                    }
                }
                
                // 处理合并单元格
                Map<String, Object> config = (Map<String, Object>) sheetData.get("config");
                if (config != null) {
                    Object mergeObj = config.get("merge");
                    if (mergeObj instanceof List) {
                        List<Map<String, Object>> merges = (List<Map<String, Object>>) mergeObj;
                        for (Map<String, Object> merge : merges) {
                            try {
                                int r = ((Number) merge.get("r")).intValue();
                                int c = ((Number) merge.get("c")).intValue();
                                int rs = ((Number) merge.get("rs")).intValue();
                                int cs = ((Number) merge.get("cs")).intValue();
                                
                                sheet.addMergedRegion(new CellRangeAddress(r, r + rs - 1, c, c + cs - 1));
                            } catch (Exception e) {
                                log.warn("处理合并单元格失败: {}", merge, e);
                            }
                        }
                    }
                }
            }
            
            // 保存文件
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * 从Luckysheet格式设置单元格值
     */
    private void setCellValueFromLuckysheet(Cell cell, Map<String, Object> v) {
        Object value = v.get("v");
        String formula = (String) v.get("f");
        
        if (formula != null && !formula.isEmpty()) {
            cell.setCellFormula(formula);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value != null) {
            cell.setCellValue(value.toString());
        }
        
        // 设置样式（简化处理）
        Map<String, Object> style = (Map<String, Object>) v.get("s");
        if (style != null) {
            CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
            
            // 字体
            Map<String, Object> ff = (Map<String, Object>) style.get("ff");
            if (ff != null) {
                Font font = cell.getSheet().getWorkbook().createFont();
                Object fs = ff.get("fs");
                if (fs instanceof Number) {
                    font.setFontHeightInPoints(((Number) fs).shortValue());
                }
                Object bl = ff.get("bl");
                if (bl instanceof Number && ((Number) bl).intValue() == 1) {
                    font.setBold(true);
                }
                cellStyle.setFont(font);
            }
            
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 获取单元格字符串值
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
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
