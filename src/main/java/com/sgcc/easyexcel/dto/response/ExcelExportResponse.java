package com.sgcc.easyexcel.dto.response;

import io.swagger.v3.oas.annotations.media.Schema;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Excel导出统一响应
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "Excel导出响应")
public class ExcelExportResponse {

    @Schema(description = "业务状态码，200表示成功，其他表示失败，具体错误码请查看文档", 
            example = "200")
    private Integer code;

    @Schema(description = "响应消息，成功时为“导出成功”，失败时为具体错误原因", 
            example = "导出成功")
    private String message;

    @Schema(description = "是否成功，true表示成功，false表示失败", 
            example = "true")
    private Boolean success;

    @Schema(description = "响应数据，成功时返回具体数据，失败时为null")
    private Object data;

    @Schema(description = "响应时间戳（毫秒）", 
            example = "1703232645000")
    private Long timestamp;

    /**
     * 成功响应（专门用于ExcelFileInfo）
     */
    public static ExcelExportResponse success(ExcelFileInfo data) {
        return ExcelExportResponse.builder()
                .code(200)
                .message("导出成功")
                .success(true)
                .data(data)
                .timestamp(System.currentTimeMillis())
                .build();
    }

    /**
     * 通用成功响应
     */
    public static ExcelExportResponse success(Object data) {
        return ExcelExportResponse.builder()
                .code(200)
                .message("操作成功")
                .success(true)
                .data(data)
                .timestamp(System.currentTimeMillis())
                .build();
    }

    /**
     * 失败响应
     */
    public static ExcelExportResponse error(Integer code, String message) {
        return ExcelExportResponse.builder()
                .code(code)
                .message(message)
                .success(false)
                .data(null)
                .timestamp(System.currentTimeMillis())
                .build();
    }
}