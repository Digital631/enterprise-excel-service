package com.sgcc.easyexcel.service;

import com.sgcc.easyexcel.dto.request.TemplateFieldMappingRequest;
import com.sgcc.easyexcel.dto.response.TemplateInfoResponse;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * 模板配置服务接口
 *
 * @author system
 * @since 2023-12-22
 */
public interface TemplateConfigService {

    /**
     * 获取所有模板列表
     *
     * @return 模板信息列表
     */
    List<TemplateInfoResponse> getAllTemplates();

    /**
     * 获取模板详细信息
     *
     * @param templateName 模板名称
     * @return 模板详细信息
     */
    TemplateInfoResponse getTemplateInfo(String templateName);

    /**
     * 上传模板文件
     *
     * @param file 模板文件
     * @return 上传后的模板信息
     * @throws IOException IO异常
     */
    TemplateInfoResponse uploadTemplate(MultipartFile file) throws IOException;

    /**
     * 删除模板文件
     *
     * @param templateName 模板名称
     * @return 是否删除成功
     */
    boolean deleteTemplate(String templateName);

    /**
     * 保存字段映射配置
     *
     * @param request 字段映射配置请求
     * @return 是否保存成功
     */
    boolean saveFieldMapping(TemplateFieldMappingRequest request);

    /**
     * 获取字段映射配置
     *
     * @param templateName 模板名称
     * @return 字段映射配置
     */
    Map<String, String> getFieldMapping(String templateName);

    /**
     * 预览模板数据
     *
     * @param templateName 模板名称
     * @param sheetIndex Sheet索引
     * @return 预览数据
     */
    List<Map<String, Object>> previewTemplateData(String templateName, int sheetIndex);

    /**
     * 验证模板文件
     *
     * @param templateName 模板名称
     * @return 验证结果
     */
    Map<String, Object> validateTemplate(String templateName);
}
