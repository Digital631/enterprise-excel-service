package com.sgcc.easyexcel.dto.request;

import io.swagger.v3.oas.annotations.media.Schema;
import jakarta.validation.constraints.NotBlank;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * 模板保存请求
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "模板保存请求")
public class TemplateSaveRequest {

    @NotBlank(message = "模板名称不能为空")
    @Schema(description = "模板名称", required = true)
    private String templateName;

    @Schema(description = "模板描述")
    private String description;

    @Schema(description = "Luckysheet数据，包含多个Sheet的配置信息")
    private List<Map<String, Object>> sheets;

    @Schema(description = "是否覆盖已存在的模板", defaultValue = "false")
    private Boolean overwrite = false;

    @Schema(description = "模板配置选项")
    private TemplateOptions options;

    /**
     * 模板配置选项
     */
    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    @Schema(description = "模板配置选项")
    public static class TemplateOptions {

        @Schema(description = "默认数据起始行", example = "2")
        private Integer defaultStartRow;

        @Schema(description = "表头行号", example = "1")
        private Integer headerRow;

        @Schema(description = "是否包含公式", defaultValue = "true")
        private Boolean includeFormulas;

        @Schema(description = "是否保留样式", defaultValue = "true")
        private Boolean preserveStyles;
    }
}
