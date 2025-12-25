package com.sgcc.easyexcel.dto.response;

import io.swagger.v3.oas.annotations.media.Schema;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

/**
 * Excel文件信息
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Schema(description = "Excel文件信息")
public class ExcelFileInfo {

    @Schema(description = "文件路径", example = "E:/exports/员工统计表_20231222153010.xlsx")
    private String filePath;

    @Schema(description = "文件名", example = "员工统计表_20231222153010.xlsx")
    private String fileName;

    @Schema(description = "文件大小（字节）", example = "102400")
    private Long fileSize;

    @Schema(description = "导出时间", example = "2023-12-22 15:30:10")
    private String exportTime;

    @Schema(description = "导出耗时（毫秒）", example = "1500")
    private Long exportDuration;

    @Schema(description = "导出耗时（秒，保留2位小数）", example = "1.50")
    private BigDecimal exportDurationSeconds;
}