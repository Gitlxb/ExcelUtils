package com.example.project.config;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * SPA 回退控制器（修复版）
 * 只拦截前端路由路径，不拦截静态资源（js/css/img 等）
 * 静态文件仍由 Spring Boot ResourceHttpRequestHandler 正常提供
 */
@Controller
public class SpaForwardController {

    /**
     * 只处理主页 + 前端路由，不拦截带后缀的静态文件
     * spring.mvc.static-path-pattern 默认是 /**，所以静态资源优先级更高
     *
     * 因为 @RequestMapping 的优先级低于 ResourceHttpRequestHandler，
     * 所以 /js/xxx.js 等会先被静态资源处理器匹配，不会走到这里。
     */
    @RequestMapping(value = {
        "/",
        "/general/**",
        "/finance/**",
        "/hr/**"
    })
    public String forwardToIndex() {
        return "forward:/index.html";
    }
}
