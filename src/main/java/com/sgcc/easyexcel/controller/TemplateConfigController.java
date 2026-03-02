package com.sgcc.easyexcel.controller;

import com.sgcc.easyexcel.dto.request.TemplateFieldMappingRequest;
import com.sgcc.easyexcel.dto.response.ExcelExportResponse;
import com.sgcc.easyexcel.dto.response.TemplateInfoResponse;
import com.sgcc.easyexcel.service.TemplateConfigService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.util.StreamUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;

/**
 * 模板配置控制器
 * 提供模板管理和配置的Web界面及API
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Controller
@RequestMapping("/template")
@Tag(name = "模板配置管理", description = "Excel模板上传、配置、管理功能")
public class TemplateConfigController {

    private final TemplateConfigService templateConfigService;

    public TemplateConfigController(TemplateConfigService templateConfigService) {
        this.templateConfigService = templateConfigService;
    }

    /**
     * 模板管理页面
     */
    @GetMapping({"", "/", "/index"})
    @Operation(summary = "模板管理页面", description = "返回模板配置的Web管理界面")
    public String index() {
        return "forward:/template/index.html";
    }

    /**
     * 获取所有模板列表（API）
     */
    @GetMapping("/api/list")
    @ResponseBody
    @Operation(summary = "获取模板列表", description = "获取所有已上传的Excel模板列表")
    public ExcelExportResponse getTemplateList() {
        try {
            List<TemplateInfoResponse> templates = templateConfigService.getAllTemplates();
            return ExcelExportResponse.success(templates);
        } catch (Exception e) {
            log.error("获取模板列表失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 获取模板详细信息（API）
     */
    @GetMapping("/api/info/{templateName}")
    @ResponseBody
    @Operation(summary = "获取模板详情", description = "获取指定模板的详细信息")
    public ExcelExportResponse getTemplateInfo(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            TemplateInfoResponse template = templateConfigService.getTemplateInfo(templateName);
            return ExcelExportResponse.success(template);
        } catch (Exception e) {
            log.error("获取模板详情失败：{}" , templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 上传模板文件（API）
     */
    @PostMapping("/api/upload")
    @ResponseBody
    @Operation(summary = "上传模板", description = "上传Excel模板文件")
    public ExcelExportResponse uploadTemplate(
            @Parameter(description = "模板文件", required = true)
            @RequestParam("file") MultipartFile file) {
        try {
            if (file.isEmpty()) {
                return ExcelExportResponse.error(400, "上传文件为空");
            }
            
            TemplateInfoResponse template = templateConfigService.uploadTemplate(file);
            return ExcelExportResponse.success(template);
        } catch (Exception e) {
            log.error("上传模板失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 删除模板文件（API）
     */
    @DeleteMapping("/api/delete/{templateName}")
    @ResponseBody
    @Operation(summary = "删除模板", description = "删除指定的模板文件")
    public ExcelExportResponse deleteTemplate(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            boolean success = templateConfigService.deleteTemplate(templateName);
            if (success) {
                return ExcelExportResponse.success("删除成功");
            } else {
                return ExcelExportResponse.error(500, "删除失败");
            }
        } catch (Exception e) {
            log.error("删除模板失败：{}" , templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 保存字段映射配置（API）
     */
    @PostMapping("/api/mapping")
    @ResponseBody
    @Operation(summary = "保存字段映射", description = "保存模板的字段映射配置")
    public ExcelExportResponse saveFieldMapping(
            @Parameter(description = "字段映射配置", required = true)
            @RequestBody TemplateFieldMappingRequest request) {
        try {
            boolean success = templateConfigService.saveFieldMapping(request);
            if (success) {
                return ExcelExportResponse.success("保存成功");
            } else {
                return ExcelExportResponse.error(500, "保存失败");
            }
        } catch (Exception e) {
            log.error("保存字段映射失败", e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 获取字段映射配置（API）
     */
    @GetMapping("/api/mapping/{templateName}")
    @ResponseBody
    @Operation(summary = "获取字段映射", description = "获取指定模板的字段映射配置")
    public ExcelExportResponse getFieldMapping(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            Map<String, String> mapping = templateConfigService.getFieldMapping(templateName);
            return ExcelExportResponse.success(mapping);
        } catch (Exception e) {
            log.error("获取字段映射失败：{}" , templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 预览模板数据（API）
     */
    @GetMapping("/api/preview/{templateName}")
    @ResponseBody
    @Operation(summary = "预览模板数据", description = "预览模板的前几行数据")
    public ExcelExportResponse previewTemplate(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName,
            @Parameter(description = "Sheet索引", example = "0")
            @RequestParam(defaultValue = "0") int sheetIndex) {
        try {
            List<Map<String, Object>> previewData = templateConfigService.previewTemplateData(templateName, sheetIndex);
            return ExcelExportResponse.success(previewData);
        } catch (Exception e) {
            log.error("预览模板数据失败：{}" , templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }

    /**
     * 验证模板（API）
     */
    @GetMapping("/api/validate/{templateName}")
    @ResponseBody
    @Operation(summary = "验证模板", description = "验证模板的有效性")
    public ExcelExportResponse validateTemplate(
            @Parameter(description = "模板名称", required = true)
            @PathVariable String templateName) {
        try {
            Map<String, Object> result = templateConfigService.validateTemplate(templateName);
            return ExcelExportResponse.success(result);
        } catch (Exception e) {
            log.error("验证模板失败：{}" , templateName, e);
            return ExcelExportResponse.error(500, e.getMessage());
        }
    }
}
