package com.sgcc.easyexcel.service;

import com.sgcc.easyexcel.dto.request.SingleSheetExportRequest;
import com.sgcc.easyexcel.dto.response.ExcelFileInfo;

/**
 * 异步Excel导出服务接口
 * 支持任务ID管理，允许停止导出操作
 *
 * @author system
 * @since 2023-12-23
 */
public interface AsyncExcelExportService {

    /**
     * 异步单Sheet导出
     *
     * @param request 导出请求
     * @param taskId 任务ID
     * @return 文件信息
     */
    ExcelFileInfo exportSingleSheetAsync(SingleSheetExportRequest request, String taskId);
}