package com.sgcc.easyexcel;

import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.core.env.Environment;

import java.net.InetAddress;
import java.net.UnknownHostException;

@Slf4j
@SpringBootApplication
public class EasyexcelApplication {

    public static void main(String[] args) throws UnknownHostException {
        ConfigurableApplicationContext application = SpringApplication.run(EasyexcelApplication.class, args);
        Environment env = application.getEnvironment();
        String ip = InetAddress.getLocalHost().getHostAddress();
        String port = env.getProperty("server.port");
        String contextPath = env.getProperty("server.servlet.context-path", "");
        
        log.info("\n----------------------------------------------------------\n\t" +
                "Application '{}' is running! Access URLs:\n\t" +
                "Knife4j文档: \thttp://localhost:{}{}/doc.html\n\t" +
                "Swagger UI: \thttp://localhost:{}{}/swagger-ui/index.html\n\t" +
                "----------------------------------------------------------",
                env.getProperty("spring.application.name"),
                port, contextPath,
                port, contextPath
        );
    }

}
