package com.sgcc.easyexcel.dto.request;

import io.swagger.v3.oas.annotations.media.Schema;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * Sheet数据配置（多Sheet用）
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "Sheet数据配置，用于多Sheet导出时配置每个Sheet的数据和映射关系")
public class SheetDataConfig {

    @NotNull(message = "Sheet索引不能为空")
    @Min(value = 0, message = "Sheet索引必须大于等于0")
    @Schema(description = "Sheet索引，从0开始，0表示第一个Sheet", 
            example = "0", 
            required = true,
            requiredMode = Schema.RequiredMode.REQUIRED)
    private Integer sheetIndex;

    @Schema(description = "Sheet名称，可选，不传则使用模板中的原始名称", 
            example = "员工信息")
    private String sheetName;

    @NotEmpty(message = "Sheet数据不能为空")
    @Schema(description = "该Sheet的数据列表", 
            required = true,
            requiredMode = Schema.RequiredMode.REQUIRED)
    private List<Map<String, Object>> data;

    @Min(value = 1, message = "起始行必须大于0")
    @Schema(description = "该Sheet的数据起始行", 
            example = "2", 
            defaultValue = "2")
    private Integer startRow = 2;

    @Schema(description = "该Sheet的字段映射关系", 
            example = "{\"姓名\":\"userName\"}")
    private Map<String, String> fieldMapping;

    @Min(value = 100, message = "分批大小不能小于100")
    @Schema(description = "该Sheet的分批大小", 
            example = "5000", 
            defaultValue = "5000")
    private Integer batchSize = 5000;
}
