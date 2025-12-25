package com.sgcc.easyexcel.config;

import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.info.Contact;
import io.swagger.v3.oas.models.info.Info;
import io.swagger.v3.oas.models.info.License;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * Knife4j和OpenAPI配置类
 *
 * @author system
 * @since 2023-12-23
 */
@Configuration
public class Knife4jConfig {

    /**
     * OpenAPI文档配置
     */
    @Bean
    public OpenAPI customOpenAPI() {
        return new OpenAPI()
                .info(new Info()
                        .title("Excel导出服务API文档")
                        .version("1.0.0")
                        .description("""
                                # 企业级通用Excel导出服务
                                
                                ## 核心功能
                                - ✅ 基于模板的Excel导出
                                - ✅ 支持大数据量（百万级）
                                - ✅ 高并发处理（50+并发）
                                - ✅ 保留模板公式和样式
                                - ✅ 灵活的字段映射
                                
                                ## 使用说明
                                1. 准备Excel模板文件（放在配置的模板目录）
                                2. 调用导出接口传入数据（Map格式）
                                3. 获取导出文件路径
                                
                                ## 性能指标
                                - 1万以内数据：响应时间 < 3秒
                                - 10万数据（大数据模式）：响应时间 < 30秒
                                - 支持50个并发请求
                                
                                ## 技术支持
                                如有问题请联系技术支持团队
                                """)
                        .contact(new Contact()
                                .name("SGCC技术团队")
                                .email("support@sgcc.com")
                                .url("https://www.sgcc.com")
                        )
                        .license(new License()
                                .name("Apache 2.0")
                                .url("https://www.apache.org/licenses/LICENSE-2.0.html")
                        )
                );
    }
}
