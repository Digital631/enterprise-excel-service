package com.sgcc.easyexcel.dto.request;

import io.swagger.v3.oas.annotations.media.Schema;
import jakarta.validation.constraints.NotBlank;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;

/**
 * 模板字段映射配置请求
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "模板字段映射配置请求")
public class TemplateFieldMappingRequest {

    @NotBlank(message = "模板名称不能为空")
    @Schema(description = "模板名称", required = true)
    private String templateName;

    @Schema(description = "字段映射关系，key为模板表头名，value为数据字段名", 
            example = "{\"姓名\":\"userName\",\"年龄\":\"age\"}")
    private Map<String, String> fieldMapping;

    @Schema(description = "匹配策略：exact/ignore-case/smart", 
            example = "ignore-case",
            defaultValue = "ignore-case")
    private String matchStrategy = "ignore-case";
}
