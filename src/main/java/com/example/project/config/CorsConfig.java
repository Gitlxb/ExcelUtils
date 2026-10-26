package com.example.project.config;

import jakarta.servlet.Filter;
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

    /**
     * 仅在配置了允许来源时才启用 CORS 校验。
     * 未配置时返回透传空过滤器（等同不启用 CORS）：
     * 同域部署下浏览器 POST 也会携带 Origin 头，若注册全局 CorsFilter 而白名单
     * 不含本站来源，所有 API 请求都会被 403 拦截（2026-10-15 事故）。
     * 注意不能返回 null——@Bean 返回 null 会被登记为 NullBean，
     * Spring Boot 按 Filter 类型收集过滤器时启动直接报错。
     */
    @Bean
    public Filter corsFilter() {
        if (allowedOrigins == null || allowedOrigins.isBlank()) {
            return (request, response, chain) -> chain.doFilter(request, response);
        }

        CorsConfiguration config = new CorsConfiguration();
        List<String> origins = Arrays.asList(allowedOrigins.split(","));
        origins.forEach(config::addAllowedOrigin);

        config.addAllowedMethod("*");
        config.addAllowedHeader("*");
        config.setAllowCredentials(true);

        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        source.registerCorsConfiguration("/**", config);
        return new CorsFilter(source);
    }
}
