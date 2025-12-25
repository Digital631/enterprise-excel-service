package com.sgcc.easyexcel.service;

import com.sgcc.easyexcel.dto.request.MultiSheetExportRequest;
import com.sgcc.easyexcel.dto.request.SingleSheetExportRequest;
import com.sgcc.easyexcel.dto.response.ExcelFileInfo;

/**
 * Excel导出服务接口
 *
 * @author system
 * @since 2023-12-22
 */
public interface ExcelExportService {

    /**
     * 单Sheet导出
     *
     * @param request 导出请求
     * @return 导出文件信息
     */
    ExcelFileInfo exportSingleSheet(SingleSheetExportRequest request);

    /**
     * 单Sheet导出（带任务ID）
     *
     * @param request 导出请求
     * @param taskId 任务ID
     * @return 导出文件信息
     */
    ExcelFileInfo exportSingleSheet(SingleSheetExportRequest request, String taskId);

    /**
     * 多Sheet导出
     *
     * @param request 导出请求
     * @return 导出文件信息
     */
    ExcelFileInfo exportMultiSheet(MultiSheetExportRequest request);
}