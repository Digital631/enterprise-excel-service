package com.sgcc.easyexcel.util;

import com.sgcc.easyexcel.dto.response.ExcelFileInfo;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.Map;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 任务管理器
 * 管理导出任务的线程引用，支持任务的启动、停止和状态查询
 *
 * @author system
 * @since 2023-12-23
 */
@Slf4j
@Component
public class TaskManager {

    /**
     * 存储导出任务的线程引用
     * Key: 任务ID
     * Value: 执行导出任务的线程
     */
    private final Map<String, Thread> exportTaskThreads = new ConcurrentHashMap<>();
    
    /**
     * 存储任务执行结果
     * Key: 任务ID
     * Value: 任务执行结果（成功时为ExcelFileInfo，失败时为异常信息）
     */
    private final Map<String, Object> taskResults = new ConcurrentHashMap<>();
    
    /**
     * 存储任务状态
     * Key: 任务ID
     * Value: 任务状态
     */
    private final Map<String, TaskStatus> taskStatuses = new ConcurrentHashMap<>();

    /**
     * 任务状态枚举
     */
    public enum TaskStatus {
        RUNNING("运行中"),
        COMPLETED("已完成"),
        FAILED("已失败"),
        STOPPED("已停止");

        private final String description;

        TaskStatus(String description) {
            this.description = description;
        }

        public String getDescription() {
            return description;
        }
    }

    /**
     * 生成唯一的任务ID
     *
     * @return 任务ID
     */
    public String generateTaskId() {
        String taskId = UUID.randomUUID().toString().replace("-", "");
        log.debug("生成任务ID：{}", taskId);
        return taskId;
    }

    /**
     * 注册导出任务
     *
     * @param taskId 任务ID
     * @param thread 执行任务的线程
     */
    public void registerExportTask(String taskId, Thread thread) {
        exportTaskThreads.put(taskId, thread);
        taskStatuses.put(taskId, TaskStatus.RUNNING);
        log.debug("注册导出任务：taskId={}, thread={}", taskId, thread.getName());
    }

    /**
     * 取消注册导出任务
     *
     * @param taskId 任务ID
     */
    public void unregisterExportTask(String taskId) {
        exportTaskThreads.remove(taskId);
        log.debug("取消注册导出任务：taskId={}", taskId);
    }

    /**
     * 停止指定的导出任务
     *
     * @param taskId 任务ID
     * @return 是否成功停止任务
     */
    public boolean stopExportTask(String taskId) {
        Thread thread = exportTaskThreads.get(taskId);
        if (thread != null) {
            log.info("停止导出任务：taskId={}", taskId);
            thread.interrupt();
            taskStatuses.put(taskId, TaskStatus.STOPPED);
            exportTaskThreads.remove(taskId); // 移除线程引用
            return true;
        }
        return false;
    }

    /**
     * 获取任务状态
     *
     * @param taskId 任务ID
     * @return 任务状态
     */
    public TaskStatus getTaskStatus(String taskId) {
        return taskStatuses.getOrDefault(taskId, null);
    }

    /**
     * 设置任务结果
     *
     * @param taskId 任务ID
     * @param result 任务结果（成功时为ExcelFileInfo，失败时为异常信息）
     * @param status 任务状态
     */
    public void setTaskResult(String taskId, Object result, TaskStatus status) {
        taskResults.put(taskId, result);
        taskStatuses.put(taskId, status);
        
        // 如果任务完成，移除线程引用
        if (status == TaskStatus.COMPLETED || status == TaskStatus.FAILED || status == TaskStatus.STOPPED) {
            exportTaskThreads.remove(taskId);
        }
        
        log.debug("设置任务结果：taskId={}, status={}", taskId, status);
    }

    /**
     * 获取任务结果
     *
     * @param taskId 任务ID
     * @return 任务结果
     */
    public Object getTaskResult(String taskId) {
        return taskResults.get(taskId);
    }

    /**
     * 获取所有活跃任务
     *
     * @return 活跃任务状态映射
     */
    public Map<String, TaskStatus> getActiveTasks() {
        return new ConcurrentHashMap<>(taskStatuses);
    }
}