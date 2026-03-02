package com.sgcc.easyexcel.dto.response;

import io.swagger.v3.oas.annotations.media.Schema;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

/**
 * 模板信息响应
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "模板信息响应")
public class TemplateInfoResponse {

    @Schema(description = "模板文件名")
    private String fileName;

    @Schema(description = "模板完整路径")
    private String filePath;

    @Schema(description = "文件大小（字节）")
    private Long fileSize;

    @Schema(description = "文件大小（格式化）")
    private String fileSizeFormatted;

    @Schema(description = "最后修改时间")
    private LocalDateTime lastModified;

    @Schema(description = "Sheet列表")
    private List<SheetInfo> sheets;

    @Schema(description = "是否已配置字段映射")
    private Boolean hasFieldMapping;

    @Schema(description = "字段映射配置")
    private Map<String, String> fieldMapping;

    /**
     * Sheet信息
     */
    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    @Schema(description = "Sheet信息")
    public static class SheetInfo {

        @Schema(description = "Sheet索引")
        private Integer sheetIndex;

        @Schema(description = "Sheet名称")
        private String sheetName;

        @Schema(description = "表头列表")
        private List<String> headers;

        @Schema(description = "数据行数")
        private Integer dataRowCount;
    }
}
