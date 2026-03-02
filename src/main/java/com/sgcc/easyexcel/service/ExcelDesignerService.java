package com.sgcc.easyexcel.service;

import com.sgcc.easyexcel.dto.request.TemplateSaveRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Excel模板设计器服务接口
 *
 * @author system
 * @since 2023-12-22
 */
public interface ExcelDesignerService {

    /**
     * 获取模板列表
     *
     * @return 模板列表
     */
    List<Map<String, Object>> getTemplateList();

    /**
     * 加载模板数据
     *
     * @param templateName 模板名称
     * @return Luckysheet格式的模板数据
     */
    Map<String, Object> loadTemplate(String templateName);

    /**
     * 保存模板
     *
     * @param request 保存请求
     * @return 保存的文件路径
     */
    String saveTemplate(TemplateSaveRequest request);

    /**
     * 导入Excel文件
     *
     * @param file Excel文件
     * @return Luckysheet格式的数据
     * @throws IOException IO异常
     */
    Map<String, Object> importExcel(MultipartFile file) throws IOException;

    /**
     * 导出Excel文件
     *
     * @param templateName 模板名称
     * @param response HTTP响应
     * @throws IOException IO异常
     */
    void exportExcel(String templateName, HttpServletResponse response) throws IOException;

    /**
     * 删除模板
     *
     * @param templateName 模板名称
     * @return 是否删除成功
     */
    boolean deleteTemplate(String templateName);

    /**
     * 提取模板中的占位符
     *
     * @param templateName 模板名称
     * @return 占位符列表
     */
    List<String> extractPlaceholders(String templateName);

    /**
     * 预览模板填充效果
     *
     * @param templateName 模板名称
     * @param sampleData 示例数据
     * @return 预览数据
     */
    Map<String, Object> previewTemplate(String templateName, Map<String, Object> sampleData);
}
