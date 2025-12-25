package com.sgcc.easyexcel.controller;

import com.sgcc.easyexcel.dto.response.ExcelExportResponse;
import com.sgcc.easyexcel.util.TaskManager;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;

/**
 * 导出任务控制器
 * 提供任务管理功能，如停止任务、查询状态等
 *
 * @author system
 * @since 2023-12-23
 */
@Slf4j
@RestController
@RequestMapping("/api/excel/task")
@Tag(name = "导出任务管理", description = "提供导出任务的管理功能，如停止任务、查询状态等")
public class ExportTaskController {

    private final TaskManager taskManager;

    public ExportTaskController(TaskManager taskManager) {
        this.taskManager = taskManager;
    }

    /**
     * 停止导出任务
     *
     * @param taskId 任务ID
     * @return 操作结果
     */
    @PostMapping("/stop/{taskId}")
    @Operation(
            summary = "停止导出任务",
            description = "停止指定ID的导出任务，中断正在执行的导出操作"
    )
    @ApiResponse(responseCode = "200", description = "操作成功")
    public ExcelExportResponse stopTask(
            @Parameter(description = "任务ID", required = true)
            @PathVariable String taskId
    ) {
        log.info("收到停止任务请求：taskId={}", taskId);
        
        boolean success = taskManager.stopExportTask(taskId);
        if (success) {
            log.info("成功停止任务：taskId={}", taskId);
            return ExcelExportResponse.success("任务已停止");
        } else {
            log.warn("停止任务失败，任务不存在：taskId={}", taskId);
            return ExcelExportResponse.error(404, "任务不存在或已完成");
        }
    }

    /**
     * 查询任务状态
     *
     * @param taskId 任务ID
     * @return 任务状态信息
     */
    @GetMapping("/status/{taskId}")
    @Operation(
            summary = "查询任务状态",
            description = "查询指定ID的导出任务状态，包括运行中、已完成、已失败、已停止等状态"
    )
    @ApiResponse(responseCode = "200", description = "查询成功")
    public ExcelExportResponse getTaskStatus(
            @Parameter(description = "任务ID", required = true)
            @PathVariable String taskId
    ) {
        log.info("收到查询任务状态请求：taskId={}", taskId);
        
        TaskManager.TaskStatus status = taskManager.getTaskStatus(taskId);
        if (status == null) {
            log.warn("查询任务状态失败，任务不存在：taskId={}", taskId);
            return ExcelExportResponse.error(404, "任务不存在");
        }
        
        // 构建状态信息
        java.util.Map<String, Object> statusInfo = new java.util.HashMap<>();
        statusInfo.put("taskId", taskId);
        statusInfo.put("status", status.name());
        statusInfo.put("statusDescription", status.getDescription());
        
        // 如果任务已完成，包含结果信息
        if (status == TaskManager.TaskStatus.COMPLETED || 
            status == TaskManager.TaskStatus.FAILED || 
            status == TaskManager.TaskStatus.STOPPED) {
            Object result = taskManager.getTaskResult(taskId);
            statusInfo.put("result", result);
        }
        
        return ExcelExportResponse.success(statusInfo);
    }

    /**
     * 获取所有活跃任务
     *
     * @return 活跃任务列表
     */
    @GetMapping("/active")
    @Operation(
            summary = "获取所有活跃任务",
            description = "获取当前所有活跃的导出任务状态"
    )
    @ApiResponse(responseCode = "200", description = "查询成功")
    public ExcelExportResponse getActiveTasks() {
        log.info("收到查询活跃任务请求");
        
        java.util.Map<String, TaskManager.TaskStatus> activeTasks = taskManager.getActiveTasks();
        return ExcelExportResponse.success(activeTasks);
    }
}