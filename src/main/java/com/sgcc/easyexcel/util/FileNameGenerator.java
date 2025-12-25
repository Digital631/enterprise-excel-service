package com.sgcc.easyexcel.util;

import com.sgcc.easyexcel.config.ExcelConfig;
import com.sgcc.easyexcel.enums.ExcelErrorCode;
import com.sgcc.easyexcel.exception.ExcelExportException;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.UUID;

/**
 * 文件名生成器
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Component
public class FileNameGenerator {

    private final ExcelConfig excelConfig;

    public FileNameGenerator(ExcelConfig excelConfig) {
        this.excelConfig = excelConfig;
    }

    /**
     * 生成导出文件名（线程安全）
     *
     * @param templateName 模板名称
     * @param customFileName 自定义文件名
     * @return 文件名
     */
    public synchronized String generateFileName(String templateName, String customFileName) {
        ExcelConfig.ExportConfig config = excelConfig.getExport();
        
        // 如果指定了自定义文件名
        if (customFileName != null && !customFileName.isEmpty()) {
            return validateAndFormatFileName(customFileName);
        }
        
        // 根据策略生成文件名
        String baseName = getBaseName(templateName);
        String strategy = config.getFilenameStrategy();
        
        return switch (strategy) {
            case "timestamp" -> generateTimestampFileName(baseName);
            case "uuid" -> generateUuidFileName(baseName);
            case "original" -> baseName + ".xlsx";
            default -> generateTimestampFileName(baseName);
        };
    }

    /**
     * 处理文件重名
     *
     * @param filePath 文件完整路径
     * @return 处理后的文件路径
     */
    public String handleDuplicateFile(String filePath) {
        File file = new File(filePath);
        
        if (!file.exists()) {
            return filePath;
        }
        
        String strategy = excelConfig.getExport().getDuplicateStrategy();
        
        return switch (strategy) {
            case "overwrite" -> {
                log.warn("文件已存在，将覆盖：{}", filePath);
                yield filePath;
            }
            case "rename" -> {
                String newPath = generateUniqueFileName(filePath);
                log.info("文件已存在，自动重命名：{} -> {}", filePath, newPath);
                yield newPath;
            }
            case "error" -> throw new ExcelExportException(
                    ExcelErrorCode.FILE_ALREADY_EXISTS,
                    "文件已存在：" + filePath
            );
            default -> generateUniqueFileName(filePath);
        };
    }

    /**
     * 验证并格式化文件名
     */
    private String validateAndFormatFileName(String fileName) {
        if (!fileName.toLowerCase().endsWith(".xlsx") && !fileName.toLowerCase().endsWith(".xls")) {
            throw new ExcelExportException(ExcelErrorCode.INVALID_FILE_NAME);
        }
        
        // 移除非法字符
        return fileName.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    /**
     * 获取基础文件名（移除扩展名）
     */
    private String getBaseName(String templateName) {
        if (templateName == null || templateName.isEmpty()) {
            return "export";
        }
        
        // 移除路径
        String baseName = new File(templateName).getName();
        
        // 移除扩展名
        int dotIndex = baseName.lastIndexOf('.');
        if (dotIndex > 0) {
            baseName = baseName.substring(0, dotIndex);
        }
        
        return baseName;
    }

    /**
     * 生成带时间戳的文件名
     */
    private String generateTimestampFileName(String baseName) {
        String format = excelConfig.getExport().getTimestampFormat();
        String timestamp = new SimpleDateFormat(format).format(new Date());
        return baseName + "_" + timestamp + ".xlsx";
    }

    /**
     * 生成带UUID的文件名
     */
    private String generateUuidFileName(String baseName) {
        String uuid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
        return baseName + "_" + uuid + ".xlsx";
    }

    /**
     * 生成唯一文件名（追加序号）
     */
    private String generateUniqueFileName(String originalPath) {
        File originalFile = new File(originalPath);
        String dir = originalFile.getParent();
        String fileName = originalFile.getName();
        
        int dotIndex = fileName.lastIndexOf('.');
        String baseName = fileName.substring(0, dotIndex);
        String extension = fileName.substring(dotIndex);
        
        int counter = 1;
        String newPath;
        
        do {
            String newFileName = baseName + "(" + counter + ")" + extension;
            newPath = dir + File.separator + newFileName;
            counter++;
        } while (new File(newPath).exists() && counter < 1000);
        
        return newPath;
    }

    /**
     * 格式化文件大小
     */
    public String formatFileSize(long size) {
        if (size < 1024) {
            return size + " B";
        } else if (size < 1024 * 1024) {
            return String.format("%.1f KB", size / 1024.0);
        } else if (size < 1024 * 1024 * 1024) {
            return String.format("%.1f MB", size / (1024.0 * 1024));
        } else {
            return String.format("%.1f GB", size / (1024.0 * 1024 * 1024));
        }
    }
}
