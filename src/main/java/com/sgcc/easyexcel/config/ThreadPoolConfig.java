package com.sgcc.easyexcel.config;

import lombok.extern.slf4j.Slf4j;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;

import java.util.concurrent.Executor;
import java.util.concurrent.ThreadPoolExecutor;

/**
 * 线程池配置（高并发支持）
 *
 * @author system
 * @since 2023-12-22
 */
@Slf4j
@Configuration
public class ThreadPoolConfig {

    private final ExcelConfig excelConfig;

    public ThreadPoolConfig(ExcelConfig excelConfig) {
        this.excelConfig = excelConfig;
    }

    /**
     * Excel导出专用线程池
     * 
     * 高并发特性：
     * 1. 核心线程数10，最大50，支持高并发导出
     * 2. 使用CallerRunsPolicy拒绝策略，防止任务丢失
     * 3. 线程名称带前缀，便于监控和问题排查
     */
    @Bean(name = "excelExportExecutor")
    public Executor excelExportExecutor() {
        ExcelConfig.ThreadPoolConfig config = excelConfig.getThreadPool();
        
        ThreadPoolTaskExecutor executor = new ThreadPoolTaskExecutor();
        
        // 核心线程数
        executor.setCorePoolSize(config.getCorePoolSize());
        
        // 最大线程数
        executor.setMaxPoolSize(config.getMaxPoolSize());
        
        // 队列容量
        executor.setQueueCapacity(config.getQueueCapacity());
        
        // 线程空闲时间
        executor.setKeepAliveSeconds(config.getKeepAliveSeconds());
        
        // 线程名称前缀
        executor.setThreadNamePrefix(config.getThreadNamePrefix());
        
        // 拒绝策略：由调用线程处理任务（防止任务丢失）
        executor.setRejectedExecutionHandler(new ThreadPoolExecutor.CallerRunsPolicy());
        
        // 等待任务完成后再关闭线程池
        executor.setWaitForTasksToCompleteOnShutdown(true);
        
        // 最长等待时间
        executor.setAwaitTerminationSeconds(60);
        
        // 初始化
        executor.initialize();
        
        log.info("Excel导出线程池初始化完成：corePoolSize={}, maxPoolSize={}, queueCapacity={}",
                config.getCorePoolSize(), config.getMaxPoolSize(), config.getQueueCapacity());
        
        return executor;
    }
}
