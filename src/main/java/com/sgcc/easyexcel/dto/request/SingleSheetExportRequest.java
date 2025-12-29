package com.sgcc.easyexcel.dto.request;

import io.swagger.v3.oas.annotations.media.Schema;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.Min;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * 单Sheet Excel导出请求参数
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "单Sheet导出请求参数")
public class SingleSheetExportRequest {

    @NotBlank(message = "模板路径不能为空")
    @Schema(description = "模板标识，支持完整路径或模板名称", 
            example = "员工统计表", 
            required = true,
            requiredMode = Schema.RequiredMode.REQUIRED)
    private String templatePath;

    @NotEmpty(message = "导出数据不能为空")
    @Schema(description = "导出数据列表，每个Map代表一行数据，key为字段名，value为字段值", 
            example = "[{\"userName\":\"张三\",\"age\":25,\"salary\":10000}]",
            required = true,
            requiredMode = Schema.RequiredMode.REQUIRED)
    private List<Map<String, Object>> data;

    @Min(value = 1, message = "起始行必须大于0")
    @Schema(description = "数据起始行，从模板的第几行开始填充数据（默认2，因为第1行通常是表头）", 
            example = "2", 
            defaultValue = "2")
    private Integer startRow = 2;

    @Schema(description = "字段映射关系，当模板表头名称与数据字段名不一致时使用。key为模板表头名，value为数据字段名", 
            example = "{\"姓名\":\"userName\",\"年龄\":\"age\",\"工资\":\"salary\"}")
    private Map<String, String> fieldMapping;

    @Schema(description = "导出文件名，不传则自动生成（格式：模板名_时间戳.xlsx）", 
            example = "2023年12月员工统计.xlsx")
    private String exportFileName;

    @Schema(description = "导出目录，不传则使用默认配置目录", 
            example = "E:/custom-exports/")
    private String exportDir;

    @Schema(description = "是否启用大数据模式，当数据量超过10万时建议启用，采用分批写入方式防止内存溢出", 
            example = "false", 
            defaultValue = "false")
    private Boolean enableBigDataMode = false;

    @Min(value = 100, message = "分批大小不能小于100")
    @Schema(description = "分批大小，大数据模式下生效，每批次写入的数据行数，建议1000-10000", 
            example = "5000", 
            defaultValue = "5000")
    private Integer batchSize = 5000;

    @Schema(description = "单个Sheet最大行数，超过此值将自动拆分到多个Sheet，仅在启用大数据模式时生效", 
            example = "50000", 
            defaultValue = "50000")
    private Integer maxRowsPerSheet;
    
    @Schema(description = "Sheet名称", 
            example = "员工统计")
    private String sheetName;
    
    @Schema(description = "占位符数据，用于填充模板中的文档级占位符，支持多种格式如{title}、{date}、${time}、${year}等", 
            example = "{\"title\":\"员工统计表\",\"date\":\"2023-12-22\",\"time\":\"14:30:00\"}")
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
