package com.sgcc.easyexcel.service.impl;

import com.sgcc.easyexcel.service.FontService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import jakarta.annotation.PostConstruct;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 字体管理服务实现
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Service
public class FontServiceImpl implements FontService {

    // 字体文件存储目录（从配置文件读取，默认使用用户主目录）
    @Value("${excel.font.directory:${user.home}/excel-service/fonts}")
    private String fontDirectory;
    
    // 支持的字体格式
    private static final List<String> SUPPORTED_EXTENSIONS = List.of(".ttf", ".otf", ".woff", ".woff2");

    @PostConstruct
    public void init() {
        // 确保字体目录存在
        File fontDir = new File(fontDirectory);
        if (!fontDir.exists()) {
            boolean created = fontDir.mkdirs();
            if (created) {
                log.info("创建字体目录：{}", fontDirectory);
            }
        }
    }

    @Override
    public List<Map<String, Object>> getFontList() {
        List<Map<String, Object>> fonts = new ArrayList<>();
        File fontDir = new File(fontDirectory);

        if (!fontDir.exists() || !fontDir.isDirectory()) {
            return fonts;
        }

        File[] files = fontDir.listFiles((dir, name) -> {
            String lowerName = name.toLowerCase();
            return SUPPORTED_EXTENSIONS.stream().anyMatch(lowerName::endsWith);
        });

        if (files != null) {
            for (File file : files) {
                Map<String, Object> fontInfo = new HashMap<>();
                String fileName = file.getName();
                String fontName = fileName.substring(0, fileName.lastIndexOf('.'));
                String ext = fileName.substring(fileName.lastIndexOf('.'));
                
                fontInfo.put("fileName", fileName);
                fontInfo.put("fontName", fontName);
                fontInfo.put("extension", ext);
                fontInfo.put("size", file.length());
                fontInfo.put("sizeFormatted", formatFileSize(file.length()));
                fontInfo.put("lastModified", LocalDateTime.ofInstant(
                        Instant.ofEpochMilli(file.lastModified()),
                        ZoneId.systemDefault()));
                
                fonts.add(fontInfo);
            }
        }

        // 按文件名排序
        fonts.sort((a, b) -> ((String) a.get("fontName")).compareToIgnoreCase((String) b.get("fontName")));
        
        return fonts;
    }

    @Override
    public Map<String, Object> uploadFont(MultipartFile file) throws IOException {
        String originalFilename = file.getOriginalFilename();
        if (originalFilename == null || originalFilename.isEmpty()) {
            throw new RuntimeException("文件名不能为空");
        }

        // 检查文件扩展名
        String lowerName = originalFilename.toLowerCase();
        boolean validExtension = SUPPORTED_EXTENSIONS.stream().anyMatch(lowerName::endsWith);
        
        if (!validExtension) {
            throw new RuntimeException("不支持的字体格式，仅支持：" + String.join(", ", SUPPORTED_EXTENSIONS));
        }

        // 确保目录存在
        Path fontDirPath = Paths.get(fontDirectory);
        if (!Files.exists(fontDirPath)) {
            Files.createDirectories(fontDirPath);
        }

        // 保存文件
        String fileName = originalFilename;
        Path targetPath = fontDirPath.resolve(fileName);
        
        // 如果文件已存在，先删除
        if (Files.exists(targetPath)) {
            Files.delete(targetPath);
        }
        
        Files.copy(file.getInputStream(), targetPath, StandardCopyOption.REPLACE_EXISTING);
        
        log.info("字体上传成功：{}", fileName);

        // 返回字体信息
        Map<String, Object> fontInfo = new HashMap<>();
        String fontName = fileName.substring(0, fileName.lastIndexOf('.'));
        String ext = fileName.substring(fileName.lastIndexOf('.'));
        
        fontInfo.put("fileName", fileName);
        fontInfo.put("fontName", fontName);
        fontInfo.put("extension", ext);
        fontInfo.put("size", file.getSize());
        fontInfo.put("sizeFormatted", formatFileSize(file.getSize()));
        fontInfo.put("message", "字体上传成功");
        
        return fontInfo;
    }

    @Override
    public Resource loadFontAsResource(String fontName) throws IOException {
        Path fontPath = Paths.get(fontDirectory).resolve(fontName);
        File fontFile = fontPath.toFile();
        
        if (!fontFile.exists() || !fontFile.isFile()) {
            throw new RuntimeException("字体文件不存在：" + fontName);
        }
        
        return new FileSystemResource(fontFile);
    }

    @Override
    public String generateFontFaceCss(String fontName) throws IOException {
        Path fontPath = Paths.get(fontDirectory).resolve(fontName);
        File fontFile = fontPath.toFile();
        
        if (!fontFile.exists() || !fontFile.isFile()) {
            throw new RuntimeException("字体文件不存在：" + fontName);
        }

        String fontFamily = fontName.substring(0, fontName.lastIndexOf('.'));
        String extension = fontName.substring(fontName.lastIndexOf('.') + 1).toLowerCase();
        
        // 根据扩展名确定format
        String format;
        switch (extension) {
            case "ttf":
                format = "truetype";
                break;
            case "otf":
                format = "opentype";
                break;
            case "woff":
                format = "woff";
                break;
            case "woff2":
                format = "woff2";
                break;
            default:
                format = "truetype";
        }

        // 构建CSS路径（使用相对路径）
        String fontUrl = "/excel-service/designer/api/fonts/download/" + fontName;

        return String.format(
            "@font-face {\n" +
            "    font-family: '%s';\n" +
            "    src: url('%s') format('%s');\n" +
            "    font-weight: normal;\n" +
            "    font-style: normal;\n" +
            "    font-display: swap;\n" +
            "}\n",
            fontFamily, fontUrl, format
        );
    }

    @Override
    public String generateAllFontsCss() {
        StringBuilder css = new StringBuilder();
        css.append("/* 自动生成的字体CSS */\n\n");
        
        List<Map<String, Object>> fonts = getFontList();
        
        for (Map<String, Object> font : fonts) {
            String fileName = (String) font.get("fileName");
            try {
                css.append(generateFontFaceCss(fileName));
                css.append("\n");
            } catch (Exception e) {
                log.warn("生成字体CSS失败：{}", fileName, e);
            }
        }
        
        return css.toString();
    }

    @Override
    public boolean deleteFont(String fontName) {
        try {
            Path fontPath = Paths.get(fontDirectory).resolve(fontName);
            File fontFile = fontPath.toFile();
            
            if (fontFile.exists()) {
                boolean deleted = fontFile.delete();
                if (deleted) {
                    log.info("字体删除成功：{}", fontName);
                }
                return deleted;
            }
            return false;
        } catch (Exception e) {
            log.error("删除字体失败：{}", fontName, e);
            return false;
        }
    }

    @Override
    public String getFontDirectory() {
        return fontDirectory;
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
