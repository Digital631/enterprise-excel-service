package com.sgcc.easyexcel.service;

import org.springframework.core.io.Resource;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * 字体管理服务接口
 *
 * @author system
 * @since 2023-12-22
 */
public interface FontService {

    /**
     * 获取字体列表
     *
     * @return 字体列表
     */
    List<Map<String, Object>> getFontList();

    /**
     * 上传字体文件
     *
     * @param file 字体文件
     * @return 字体信息
     * @throws IOException IO异常
     */
    Map<String, Object> uploadFont(MultipartFile file) throws IOException;

    /**
     * 加载字体文件为Resource
     *
     * @param fontName 字体文件名
     * @return 字体资源
     * @throws IOException IO异常
     */
    Resource loadFontAsResource(String fontName) throws IOException;

    /**
     * 生成字体CSS
     *
     * @param fontName 字体文件名
     * @return CSS样式
     * @throws IOException IO异常
     */
    String generateFontFaceCss(String fontName) throws IOException;

    /**
     * 生成所有字体的CSS
     *
     * @return CSS样式
     */
    String generateAllFontsCss();

    /**
     * 删除字体文件
     *
     * @param fontName 字体文件名
     * @return 是否删除成功
     */
    boolean deleteFont(String fontName);

    /**
     * 获取字体存储目录
     *
     * @return 字体目录路径
     */
    String getFontDirectory();
}
