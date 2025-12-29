package com.sgcc.easyexcel.dto.request;

import io.swagger.v3.oas.annotations.media.Schema;
import jakarta.validation.Valid;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotEmpty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * 多Sheet Excel导出请求参数（预留）
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "多Sheet导出请求参数")
public class MultiSheetExportRequest {

    @NotBlank(message = "模板路径不能为空")
    @Schema(description = "模板标识，支持完整路径或模板名称", 
            example = "综合报表", 
            required = true,
            requiredMode = Schema.RequiredMode.REQUIRED)
    private String templatePath;

    @Valid
    @NotEmpty(message = "Sheet数据列表不能为空")
    @Schema(description = "多个Sheet的数据配置列表，每个元素对应一个Sheet的配置", 
            required = true,
            requiredMode = Schema.RequiredMode.REQUIRED)
    private List<SheetDataConfig> sheetDataList;

    @Schema(description = "导出文件名，不传则自动生成", 
            example = "综合报表_202312.xlsx")
    private String exportFileName;

    @Schema(description = "导出目录，不传则使用默认配置目录", 
            example = "E:/reports/")
    private String exportDir;

    @Schema(description = "是否启用大数据模式", 
            example = "false", 
            defaultValue = "false")
    private Boolean enableBigDataMode = false;
    
    @Schema(description = "占位符数据，用于填充模板中的文档级占位符，支持多种格式如{title}、{date}、${time}、${year}等", 
            example = "{\"title\":\"综合报表\",\"date\":\"2023-12-22\",\"time\":\"14:30:00\"}")
    private Map<String, Object> placeholders;
    
    @Schema(description = "设置导出的Excel文件为只读模式", 
            example = "true")
    private Boolean readOnly = false;
    
    @Schema(description = "是否对Excel文件进行密码加密", 
            example = "true")
    private Boolean enableEncryption = false;
    
    @Schema(description = "Excel文件加密密码，当enableEncryption为true时生效", 
            example = "123456")
    private String encryptionPassword;}
