package com.example.project.config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.cors.UrlBasedCorsConfigurationSource;
import org.springframework.web.filter.CorsFilter;

import java.util.Arrays;
import java.util.List;

/**
 * CORS 跨域配置
 * 前后端同域部署时此配置自动失效（安全），分离部署时通过环境变量 CORS_ALLOWED_ORIGINS 控制
 */
@Configuration
public class CorsConfig {

    @Value("${cors.allowed-origins:}")
    private String allowedOrigins;

    @Bean
    public CorsFilter corsFilter() {
        CorsConfiguration config = new CorsConfiguration();

        // 仅当配置了允许来源时才启用 CORS（前后端同域部署时无需设置）
        if (allowedOrigins != null && !allowedOrigins.isBlank()) {
            List<String> origins = Arrays.asList(allowedOrigins.split(","));
            origins.forEach(config::addAllowedOrigin);
        } else {
            // 默认不启用跨域（同域部署），仅保留本地开发回退
            config.addAllowedOrigin("http://localhost:8081");
        }

        config.addAllowedMethod("*");
        config.addAllowedHeader("*");
        config.setAllowCredentials(true);

        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        source.registerCorsConfiguration("/**", config);
        return new CorsFilter(source);
    }
}
