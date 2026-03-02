package com.sgcc.easyexcel.controller;

import com.sgcc.easyexcel.dto.request.TemplateSaveRequest;
import com.sgcc.easyexcel.dto.response.ExcelExportResponse;
import com.sgcc.easyexcel.service.ExcelDesignerService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.StreamUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * Excel模板设计器控制器
 * 提供基于Luckysheet的Excel在线设计功能
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@RestController
@RequestMapping("/designer")
@Tag(name = "Excel模板设计器", description = "在线Excel模板设计与编辑功能")
public class ExcelDesignerController {

    private final ExcelDesignerService designerService;

    public ExcelDesignerController(ExcelDesignerService designerService) {
        this.designerService = designerService;
    }

    /**
     * 模板设计器页面
     */
    @GetMapping({"", "/", "/index"})
    @Operation(summary = "模板设计器页面", description = "返回Excel在线设计器页面")
    public String index() {
        return "forward:/designer/index.html";
    }

    /**
     * 获取所有模板列表
     */
    @GetMapping("/api/templates")
    @Operation(summary = "获取模板列表", description = "获取所有已保存的模板列表")
    public ExcelExportResponse getTemplateList() {
        try {
            List<Map<String, Object>> templates = designerService.getTemplateList();
            return ExcelExportResponse.success(templates);
        } catch (Exception e) {
            log.error("获取模板列表失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 加载模板数据（Luckysheet格式）
     */
    @GetMapping("/api/template/{templateName}")
    @Operation(summary = "加载模板", description = "加载指定模板的Luckysheet数据")
    public ExcelExportResponse loadTemplate(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            Map<String, Object> templateData = designerService.loadTemplate(templateName);
            return ExcelExportResponse.success(templateData);
        } catch (Exception e) {
            log.error("加载模板失败：{}", templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 保存模板（Luckysheet格式）
     */
    @PostMapping("/api/template/save")
    @Operation(summary = "保存模板", description = "保存Luckysheet模板数据并导出为Excel文件")
    public ExcelExportResponse saveTemplate(
            @Parameter(description = "模板保存请求", required = true)
            @RequestBody TemplateSaveRequest request) {
        try {
            String filePath = designerService.saveTemplate(request);
            return ExcelExportResponse.success(Map.of(
                "message", "保存成功",
                "filePath", filePath,
                "templateName", request.getTemplateName()
            ));
        } catch (Exception e) {
            log.error("保存模板失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 导入Excel文件并转换为Luckysheet格式
     */
    @PostMapping("/api/template/import")
    @Operation(summary = "导入Excel", description = "将Excel文件导入为Luckysheet可编辑格式")
    public ExcelExportResponse importExcel(
            @Parameter(description = "Excel文件", required = true)
            @RequestParam("file") MultipartFile file) {
        try {
            if (file.isEmpty()) {
                return ExcelExportResponse.error(400, "上传文件为空");
            }
            
            Map<String, Object> luckysheetData = designerService.importExcel(file);
            return ExcelExportResponse.success(luckysheetData);
        } catch (Exception e) {
            log.error("导入Excel失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 导出Excel文件
     */
    @GetMapping("/api/template/export/{templateName}")
    @Operation(summary = "导出Excel", description = "将模板导出为Excel文件下载")
    public void exportExcel(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName,
            HttpServletResponse response) {
        try {
            designerService.exportExcel(templateName, response);
        } catch (Exception e) {
            log.error("导出Excel失败：{}", templateName, e);
            response.setStatus(500);
        }
    }

    /**
     * 删除模板
     */
    @DeleteMapping("/api/template/{templateName}")
    @Operation(summary = "删除模板", description = "删除指定的模板文件")
    public ExcelExportResponse deleteTemplate(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            boolean success = designerService.deleteTemplate(templateName);
            if (success) {
                return ExcelExportResponse.success("删除成功");
            } else {
                return ExcelExportResponse.error(500, "删除失败");
            }
        } catch (Exception e) {
            log.error("删除模板失败：{}", templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 获取模板占位符列表
     */
    @GetMapping("/api/template/placeholders/{templateName}")
    @Operation(summary = "获取占位符", description = "获取模板中定义的所有占位符")
    public ExcelExportResponse getPlaceholders(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            List<String> placeholders = designerService.extractPlaceholders(templateName);
            return ExcelExportResponse.success(placeholders);
        } catch (Exception e) {
            log.error("获取占位符失败：{}", templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 预览模板填充效果
     */
    @PostMapping("/api/template/preview")
    @Operation(summary = "预览模板", description = "使用示例数据预览模板填充效果")
    public ExcelExportResponse previewTemplate(
            @Parameter(description = "模板名称", required = true)
            @RequestParam String templateName,
            @Parameter(description = "示例数据")
            @RequestBody Map<String, Object> sampleData) {
        try {
            Map<String, Object> previewData = designerService.previewTemplate(templateName, sampleData);
            return ExcelExportResponse.success(previewData);
        } catch (Exception e) {
            log.error("预览模板失败：{}", templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }
}
