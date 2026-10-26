package com.example.project.controller;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.multipart.MaxUploadSizeExceededException;

import java.util.HashMap;
import java.util.Map;

/**
 * 全局异常处理器
 * 确保所有异常都返回标准 JSON 格式，避免前端出现 "Network Error"
 */
@ControllerAdvice
public class GlobalExceptionHandler {

    @Value("${spring.servlet.multipart.max-file-size:200MB}")
    private String maxFileSize;

    @Value("${spring.servlet.multipart.max-request-size:300MB}")
    private String maxRequestSize;

    /**
     * 上传超限：给出明确可执行的提示。
     *
     * 注意：该异常并不只表示"体积超限"——Spring 会把 Tomcat 分片数量超限
     * （Connector.maxPartCount，默认 50）抛出的 IllegalStateException 也包装成本异常，
     * 因此文案里同时提示"文件数量"。分片数已在 {@code TomcatMultipartConfig} 放宽到 2000。
     */
    @ExceptionHandler(MaxUploadSizeExceededException.class)
    public ResponseEntity<Map<String, Object>> handleMaxUpload(MaxUploadSizeExceededException e) {
        Map<String, Object> result = new HashMap<>();
        result.put("success", false);
        result.put("message", "上传内容超过服务器限制（单文件 " + maxFileSize + "、单次请求 " + maxRequestSize
                + "）。请减少单批文件数量或分批处理；若文件总量不大仍报此错，"
                + "通常是后端未重启、分片数上限（server.tomcat.max-part-count）未生效。");
        e.printStackTrace();
        return ResponseEntity.ok(result);
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<Map<String, Object>> handleException(Exception e) {
        Map<String, Object> result = new HashMap<>();
        result.put("success", false);
        String msg = e.getMessage();
        if (msg == null || msg.isEmpty()) {
            msg = e.getClass().getSimpleName();
        }
        result.put("message", "服务器错误：" + msg);
        e.printStackTrace();
        return ResponseEntity.ok(result);
    }

    @ExceptionHandler(Throwable.class)
    public ResponseEntity<Map<String, Object>> handleThrowable(Throwable t) {
        Map<String, Object> result = new HashMap<>();
        result.put("success", false);
        String msg = t.getMessage();
        if (msg == null || msg.isEmpty()) {
            msg = t.getClass().getSimpleName();
        }
        result.put("message", "服务器错误：" + msg);
        t.printStackTrace();
        return ResponseEntity.ok(result);
    }
}
