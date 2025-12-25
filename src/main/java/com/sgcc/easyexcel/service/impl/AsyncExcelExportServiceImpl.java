package com.sgcc.easyexcel.service.impl;

import com.sgcc.easyexcel.dto.request.SingleSheetExportRequest;
import com.sgcc.easyexcel.dto.response.ExcelFileInfo;
import com.sgcc.easyexcel.enums.ExcelErrorCode;
import com.sgcc.easyexcel.exception.ExcelExportException;
import com.sgcc.easyexcel.service.AsyncExcelExportService;
import com.sgcc.easyexcel.service.ExcelExportService;
import com.sgcc.easyexcel.util.TaskManager;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;

/**
 * 异步Excel导出服务实现
 * 支持任务ID管理，允许停止导出操作
 *
 * @author system
 * @since 2023-12-23
 */
@Slf4j
@Service
public class AsyncExcelExportServiceImpl implements AsyncExcelExportService {

    private final ExcelExportService excelExportService;
    private final TaskManager taskManager;

    public AsyncExcelExportServiceImpl(ExcelExportService excelExportService, TaskManager taskManager) {
        this.excelExportService = excelExportService;
        this.taskManager = taskManager;
    }

    @Override
    public ExcelFileInfo exportSingleSheetAsync(SingleSheetExportRequest request, String taskId) {
        log.info("开始异步导出任务：taskId={}, template={}, dataSize={}", 
                taskId, request.getTemplatePath(), request.getData().size());

        try {
            // 启动异步导出任务
            CompletableFuture<ExcelFileInfo> future = CompletableFuture.supplyAsync(() -> {
                Thread currentThread = Thread.currentThread();
                taskManager.registerExportTask(taskId, currentThread);

                try {
                    log.info("执行异步导出任务：taskId={}", taskId);
                    return excelExportService.exportSingleSheet(request, taskId);
                } finally {
                    taskManager.unregisterExportTask(taskId);
                    log.info("异步导出任务完成：taskId={}", taskId);
                }
            });

            // 等待结果或被中断
            return future.get();

        } catch (InterruptedException e) {
            log.warn("导出任务被中断：taskId={}", taskId);
            Thread.currentThread().interrupt(); // 重新设置中断状态
            throw new ExcelExportException(ExcelErrorCode.THREAD_INTERRUPTED, "导出任务被中断", e);
        } catch (ExecutionException e) {
            log.error("导出任务执行异常：taskId={}", taskId, e);
            Throwable cause = e.getCause();
            if (cause instanceof ExcelExportException) {
                throw (ExcelExportException) cause;
            }
            throw new ExcelExportException(ExcelErrorCode.SYSTEM_ERROR, "导出任务执行失败: " + cause.getMessage(), cause);
        } catch (Exception e) {
            log.error("导出任务异常：taskId={}", taskId, e);
            throw new ExcelExportException(ExcelErrorCode.SYSTEM_ERROR, "导出任务失败: " + e.getMessage(), e);
        }
    }
}