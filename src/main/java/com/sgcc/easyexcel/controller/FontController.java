package com.sgcc.easyexcel.controller;

import com.sgcc.easyexcel.dto.response.ExcelExportResponse;
import com.sgcc.easyexcel.service.FontService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

/**
 * 字体管理控制器
 * 提供字体上传、下载、管理功能
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@RestController
@RequestMapping("/designer/api/fonts")
@Tag(name = "字体管理", description = "自定义字体上传、下载、管理")
public class FontController {

    private final FontService fontService;

    public FontController(FontService fontService) {
        this.fontService = fontService;
    }

    /**
     * 获取所有字体列表
     */
    @GetMapping("/list")
    @Operation(summary = "获取字体列表", description = "获取所有已上传的自定义字体列表")
    public ExcelExportResponse getFontList() {
        try {
            List<Map<String, Object>> fonts = fontService.getFontList();
            return ExcelExportResponse.success(fonts);
        } catch (Exception e) {
            log.error("获取字体列表失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 上传字体文件
     */
    @PostMapping("/upload")
    @Operation(summary = "上传字体", description = "上传自定义字体文件（ttf, otf, woff, woff2）")
    public ExcelExportResponse uploadFont(
            @Parameter(description = "字体文件", required = true)
            @RequestParam("file") MultipartFile file) {
        try {
            if (file.isEmpty()) {
                return ExcelExportResponse.error(400, "上传文件为空");
            }
            
            Map<String, Object> fontInfo = fontService.uploadFont(file);
            return ExcelExportResponse.success(fontInfo);
        } catch (Exception e) {
            log.error("上传字体失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 下载字体文件
     */
    @GetMapping("/download/{fontName}")
    @Operation(summary = "下载字体", description = "下载指定字体文件")
    public ResponseEntity<Resource> downloadFont(
            @Parameter(description = "字体文件名", required = true)
            @PathVariable String fontName) {
        try {
            // 解码URL编码的文件名
            String decodedFontName = URLDecoder.decode(fontName, StandardCharsets.UTF_8);
            Resource resource = fontService.loadFontAsResource(decodedFontName);
            
            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .header(HttpHeaders.CONTENT_DISPOSITION, 
                            "attachment; filename=\"" + decodedFontName + "\"")
                    .body(resource);
        } catch (Exception e) {
            log.error("下载字体失败：{}", fontName, e);
            return ResponseEntity.notFound().build();
        }
    }

    /**
     * 获取字体CSS（用于页面加载）
     */
    @GetMapping("/css/{fontName}")
    @Operation(summary = "获取字体CSS", description = "获取字体文件的CSS样式定义")
    public ResponseEntity<String> getFontCss(
            @Parameter(description = "字体文件名", required = true)
            @PathVariable String fontName) {
        try {
            String css = fontService.generateFontFaceCss(fontName);
            
            return ResponseEntity.ok()
                    .contentType(MediaType.valueOf("text/css"))
                    .body(css);
        } catch (Exception e) {
            log.error("获取字体CSS失败：{}", fontName, e);
            return ResponseEntity.notFound().build();
        }
    }

    /**
     * 删除字体文件
     */
    @DeleteMapping("/delete/{fontName}")
    @Operation(summary = "删除字体", description = "删除指定的字体文件")
    public ExcelExportResponse deleteFont(
            @Parameter(description = "字体文件名", required = true)
            @PathVariable String fontName) {
        try {
            // 解码URL编码的文件名
            String decodedFontName = URLDecoder.decode(fontName, StandardCharsets.UTF_8);
            boolean success = fontService.deleteFont(decodedFontName);
            if (success) {
                return ExcelExportResponse.success("删除成功");
            } else {
                return ExcelExportResponse.error(500, "删除失败");
            }
        } catch (Exception e) {
            log.error("删除字体失败：{}", fontName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 获取所有字体的CSS（批量加载）
     */
    @GetMapping("/css-all")
    @Operation(summary = "获取所有字体CSS", description = "获取所有字体文件的CSS样式定义")
    public ResponseEntity<String> getAllFontsCss() {
        try {
            String css = fontService.generateAllFontsCss();
            
            return ResponseEntity.ok()
                    .contentType(MediaType.valueOf("text/css"))
                    .body(css);
        } catch (Exception e) {
            log.error("获取所有字体CSS失败", e);
            return ResponseEntity.ok().body("/* 加载字体CSS失败 */");
        }
    }
}
