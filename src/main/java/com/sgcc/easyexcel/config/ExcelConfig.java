package com.sgcc.easyexcel.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

import java.util.List;

/**
 * Excel配置类
 *
 * @author system
 * @since 2023-12-22
 */
@Data
@Configuration
@ConfigurationProperties(prefix = "excel")
public class ExcelConfig {

    /**
     * 模板配置
     */
    private TemplateConfig template = new TemplateConfig();

    /**
     * 导出配置
     */
    private ExportConfig export = new ExportConfig();

    /**
     * 填充配置
     */
    private FillConfig fill = new FillConfig();

    /**
     * 大数据配置
     */
    private BigDataConfig bigData = new BigDataConfig();

    /**
     * 监控配置
     */
    private MonitorConfig monitor = new MonitorConfig();

    /**
     * 线程池配置
     */
    private ThreadPoolConfig threadPool = new ThreadPoolConfig();

    /**
     * 获取大数据阈值
     */
    public Integer getBigDataThreshold() {
        return bigData != null ? bigData.getThreshold() : 50000; // 默认值
    }

    /**
     * 获取默认批次大小
     */
    public Integer getDefaultBatchSize() {
        return bigData != null ? bigData.getDefaultBatchSize() : 5000; // 默认值
    }

    /**
     * 获取模板路径
     */
    public String getTemplatePath() {
        return template != null ? template.getDefaultDir() : "E:\\00_WorkSpace\\DK\\easyexcel\\excel-templates";
    }

    /**
     * 获取输出路径
     */
    public String getOutputPath() {
        return export != null ? export.getDefaultDir() : "E:\\00_WorkSpace\\DK\\easyexcel\\excel-exports";
    }

    @Data
    public static class TemplateConfig {
        private String defaultDir;
        private List<String> allowedExtensions;
        private CacheConfig cache = new CacheConfig();

        @Data
        public static class CacheConfig {
            private Boolean enabled = true;
            private Integer maxSize = 50;
            private Integer expireMinutes = 30;
        }
    }

    @Data
    public static class ExportConfig {
        private String defaultDir;
        private String filenameStrategy;
        private String timestampFormat;
        private Boolean autoCreateDir;
        private String duplicateStrategy;
        private String sheetNamePrefix;
    }

    @Data
    public static class FillConfig {
        private Integer defaultStartRow;
        private Integer headerRow;
        private String fieldMatchStrategy;
    }

    @Data
    public static class BigDataConfig {
        private Integer defaultBatchSize;
        private Integer threshold;
        private Integer maxRows;
        private Integer maxRowsPerSheet;
    }

    @Data
    public static class MonitorConfig {
        private Boolean enabled;
        private Long slowThreshold;
        private Boolean detailLog;
    }

    @Data
    public static class ThreadPoolConfig {
        private Integer corePoolSize;
        private Integer maxPoolSize;
        private Integer queueCapacity;
        private Integer keepAliveSeconds;
        private String threadNamePrefix;
    }
}