package com.sgcc.easyexcel.controller;

import com.sgcc.easyexcel.dto.request.MultiSheetExportRequest;
import com.sgcc.easyexcel.dto.request.SingleSheetExportRequest;
import com.sgcc.easyexcel.dto.response.ExcelExportResponse;
import com.sgcc.easyexcel.dto.response.ExcelFileInfo;
import com.sgcc.easyexcel.enums.ExcelErrorCode;
import com.sgcc.easyexcel.exception.ExcelExportException;
import com.sgcc.easyexcel.service.AsyncExcelExportService;
import com.sgcc.easyexcel.service.ExcelExportService;
import com.sgcc.easyexcel.util.TaskManager;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.validation.Valid;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;

import java.util.concurrent.CompletableFuture;

/**
 * Excel导出控制器
 * 提供单Sheet和多Sheet的Excel导出功能
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Tag(name = "Excel导出服务", description = "企业级通用Excel导出服务，支持大数据量处理、模板导出、多Sheet导出等功能")
@RestController
@RequestMapping("/excel-service/api/excel")
public class ExcelExportController {

    private final ExcelExportService excelExportService;
    private final AsyncExcelExportService asyncExcelExportService;
    private final TaskManager taskManager;

    public ExcelExportController(ExcelExportService excelExportService, 
                                AsyncExcelExportService asyncExcelExportService,
                                TaskManager taskManager) {
        this.excelExportService = excelExportService;
        this.asyncExcelExportService = asyncExcelExportService;
        this.taskManager = taskManager;
    }

    /**
     * 单Sheet Excel导出
     *
     * <p>功能特性：</p>
     * <ul>
     *     <li>支持模板路径或模板名称</li>
     *     <li>支持字段映射配置</li>
     *     <li>支持大数据模式（10万+数据）</li>
     *     <li>保留模板公式自动计算</li>
     *     <li>高并发线程池处理</li>
     * </ul>
     *
     * <p>性能说明：</p>
     * <ul>
     *     <li>1万以内数据：响应时间 < 3秒</li>
     *     <li>10万数据（大数据模式）：响应时间 < 30秒</li>
     *     <li>支持50个并发请求</li>
     * </ul>
     *
     * <p>使用示例：</p>
     * <pre>
     * {
     *   "templatePath": "员工统计表",
     *   "data": [
     *     {"userName": "张三", "age": 25, "salary": 10000},
     *     {"userName": "李四", "age": 30, "salary": 12000}
     *   ],
     *   "startRow": 2,
     *   "fieldMapping": {
     *     "姓名": "userName",
     *     "年龄": "age",
     *     "工资": "salary"
     *   },
     *   "enableBigDataMode": false,
     *   "batchSize": 5000
     * }
     * </pre>
     *
     * @param request 导出请求参数
     * @return 导出结果（包含文件路径、大小、耗时等信息）
     */
    @PostMapping("/export/single")
    @Operation(
            summary = "单Sheet导出",
            description = "基于模板导出单个Sheet的Excel文件，支持大数据量处理和灵活的字段映射。" +
                    "当数据量超过1万条时，建议启用大数据模式以防止内存溢出。"
    )
    @ApiResponse(
            responseCode = "200",
            description = "导出成功",
            content = @Content(
                    mediaType = "application/json",
                    schema = @Schema(implementation = ExcelExportResponse.class)
            )
    )
    @ApiResponse(
            responseCode = "400",
            description = "参数错误",
            content = @Content(
                    mediaType = "application/json",
                    schema = @Schema(implementation = ExcelExportResponse.class)
            )
    )
    @ApiResponse(
            responseCode = "500",
            description = "系统错误",
            content = @Content(
                    mediaType = "application/json",
                    schema = @Schema(implementation = ExcelExportResponse.class)
            )
    )
    public ExcelExportResponse exportSingleSheet(
            @Parameter(description = "单Sheet导出请求参数", required = true)
            @Valid @RequestBody SingleSheetExportRequest request
    ) {
        log.info("收到单Sheet导出请求：template={}, dataSize={}, bigDataMode={}",
                request.getTemplatePath(),
                request.getData() != null ? request.getData().size() : 0,
                request.getEnableBigDataMode());

        try {
            // 检查数据量，如果超过1万条但未启用大数据模式，直接返回错误
            int dataSize = request.getData() != null ? request.getData().size() : 0;
            if (!request.getEnableBigDataMode() && dataSize > 10000) {
                String errorMessage = String.format("数据量为%d条，超过10000条，请将enableBigDataMode设置为true后重新请求", dataSize);
                return ExcelExportResponse.error(
                        ExcelErrorCode.DATA_TOO_LARGE.getCode(),
                        errorMessage
                );
            }

            ExcelFileInfo fileInfo = excelExportService.exportSingleSheet(request);
            return ExcelExportResponse.success(fileInfo);

        } catch (ExcelExportException e) {
            log.error("单Sheet导出失败", e);
            return ExcelExportResponse.error(e.getErrorCode().getCode(), e.getDetailMessage());
        } catch (Exception e) {
            log.error("单Sheet导出失败", e);
            return ExcelExportResponse.error(ExcelErrorCode.SYSTEM_ERROR.getCode(), e.getMessage());
        }
    }

    /**
     * 单Sheet Excel导出（异步）
     *
     * <p>功能特性：</p>
     * <ul>
     *     <li>异步执行导出任务</li>
     *     <li>支持任务进度查询</li>
     *     <li>支持任务取消</li>
     *     <li>适合长时间运行的导出任务</li>
     * </ul>
     *
     * @param request 导出请求参数
     * @return 导出结果（包含任务ID）
     */
    @PostMapping("/export/single/async")
    @Operation(
            summary = "单Sheet异步导出",
            description = "基于模板的异步导出功能，适合大数据量或长时间运行的导出任务。" +
                    "返回任务ID，可通过任务ID查询导出进度和结果。"
    )
    @ApiResponse(
            responseCode = "200",
            description = "异步导出任务提交成功",
            content = @Content(
                    mediaType = "application/json",
                    schema = @Schema(implementation = ExcelExportResponse.class)
            )
    )
    public ExcelExportResponse exportSingleSheetWithTask(
            @Parameter(description = "单Sheet导出请求参数", required = true)
            @Valid @RequestBody SingleSheetExportRequest request
    ) {
        log.info("收到带任务ID的单Sheet导出请求：template={}, dataSize={}, bigDataMode={}",
                request.getTemplatePath(),
                request.getData() != null ? request.getData().size() : 0,
                request.getEnableBigDataMode());

        try {
            // 检查数据量，如果超过1万条但未启用大数据模式，直接返回错误
            int dataSize = request.getData() != null ? request.getData().size() : 0;
            if (!request.getEnableBigDataMode() && dataSize > 10000) {
                String errorMessage = String.format("数据量为%d条，超过10000条，请将enableBigDataMode设置为true后重新请求", dataSize);
                return ExcelExportResponse.error(
                        ExcelErrorCode.DATA_TOO_LARGE.getCode(),
                        errorMessage
                );
            }

            // 生成任务ID
            String taskId = taskManager.generateTaskId();
            log.info("生成导出任务ID：{}", taskId);

            // 在新线程中执行导出任务
            CompletableFuture.runAsync(() -> {
                Thread currentThread = Thread.currentThread();
                taskManager.registerExportTask(taskId, currentThread);

                try {
                    ExcelFileInfo fileInfo = asyncExcelExportService.exportSingleSheetAsync(request, taskId);
                    log.info("异步导出任务完成：taskId={}", taskId);
                } catch (Exception e) {
                    log.error("异步导出任务失败：taskId={}", taskId, e);
                    taskManager.setTaskResult(taskId, e.getMessage(), TaskManager.TaskStatus.FAILED);
                } finally {
                    taskManager.unregisterExportTask(taskId);
                }
            });

            // 返回任务ID
            return ExcelExportResponse.success(taskId);

        } catch (Exception e) {
            log.error("异步导出任务提交失败", e);
            return ExcelExportResponse.error(ExcelErrorCode.SYSTEM_ERROR.getCode(), e.getMessage());
        }
    }

    /**
     * 查询异步导出任务状态
     *
     * @param taskId 任务ID
     * @return 任务状态信息
     */
    @GetMapping("/task/{taskId}")
    @Operation(
            summary = "查询任务状态",
            description = "查询异步导出任务的执行状态和进度信息。"
    )
    public ExcelExportResponse getTaskStatus(@PathVariable String taskId) {
        log.info("查询任务状态：taskId={}", taskId);
        try {
            Object taskResult = taskManager.getTaskResult(taskId);
            TaskManager.TaskStatus taskStatus = taskManager.getTaskStatus(taskId);
            
            return ExcelExportResponse.success(taskResult);
        } catch (Exception e) {
            log.error("查询任务状态失败：taskId={}", taskId, e);
            return ExcelExportResponse.error(ExcelErrorCode.SYSTEM_ERROR.getCode(), e.getMessage());
        }
    }

    /**
     * 停止异步导出任务
     *
     * @param taskId 任务ID
     * @return 操作结果
     */
    @DeleteMapping("/task/{taskId}")
    @Operation(
            summary = "停止任务",
            description = "停止正在执行的异步导出任务。"
    )
    public ExcelExportResponse stopTask(@PathVariable String taskId) {
        log.info("停止任务：taskId={}", taskId);
        try {
            boolean stopped = taskManager.stopTask(taskId);
            if (stopped) {
                return ExcelExportResponse.success("任务已停止");
            } else {
                return ExcelExportResponse.error(ExcelErrorCode.TASK_NOT_FOUND.getCode(), "任务不存在或已结束");
            }
        } catch (Exception e) {
            log.error("停止任务失败：taskId={}", taskId, e);
            return ExcelExportResponse.error(ExcelErrorCode.SYSTEM_ERROR.getCode(), e.getMessage());
        }
    }

    /**
     * 多Sheet Excel导出（预留）
     *
     * <p>功能特性：</p>
     * <ul>
     *     <li>支持一个Excel文件包含多个Sheet</li>
     *     <li>每个Sheet独立配置数据和映射</li>
     *     <li>支持大数据模式</li>
     * </ul>
     *
     * <p>使用场景：</p>
     * <ul>
     *     <li>综合报表（包含多个统计维度）</li>
     *     <li>多部门数据汇总</li>
     *     <li>分类数据导出</li>
     * </ul>
     *
     * @param request 多Sheet导出请求
     * @return 导出结果
     */
    @PostMapping("/export/multi")
    @Operation(
            summary = "多Sheet导出（预留）",
            description = "基于模板导出包含多个Sheet的Excel文件，适用于综合报表等场景。" +
                    "每个Sheet可以独立配置数据源和字段映射。"
    )
    @ApiResponses({
            @ApiResponse(responseCode = "200", description = "导出成功"),
            @ApiResponse(responseCode = "400", description = "参数错误"),
            @ApiResponse(responseCode = "500", description = "功能暂未实现")
    })
    public ExcelExportResponse exportMultiSheet(
            @Parameter(description = "多Sheet导出请求参数", required = true)
            @Valid @RequestBody MultiSheetExportRequest request
    ) {
        log.info("收到多Sheet导出请求：template={}, sheetCount={}",
                request.getTemplatePath(),
                request.getSheetDataList() != null ? request.getSheetDataList().size() : 0);

        try {
            ExcelFileInfo fileInfo = excelExportService.exportMultiSheet(request);
            return ExcelExportResponse.success(fileInfo);
        } catch (ExcelExportException e) {
            log.error("多Sheet导出失败", e);
            return ExcelExportResponse.error(e.getErrorCode().getCode(), e.getDetailMessage());
        } catch (Exception e) {
            log.error("多Sheet导出失败", e);
            return ExcelExportResponse.error(ExcelErrorCode.SYSTEM_ERROR.getCode(), e.getMessage());
        }
    }
}