package com.example.project.controller.general;

import com.example.project.task.FileTaskWork;
import com.example.project.task.MemoryMultipartFile;
import com.example.project.task.TaskService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

/**
 * General 模块共享基类（不注册为 Controller，仅提供工具方法）
 */
public abstract class GeneralBaseController {

    @Value("${file.upload-dir:./uploads}")
    protected String uploadDir;

    @Autowired
    protected TaskService taskService;

    /**
     * 异步提交便捷方法（单文件型工具统一入口）：
     * 请求线程内把 MultipartFile 固化为内存副本（异步安全）→ 提交后台线程池 →
     * 立即返回 {success,taskId}。work 内抛任意异常由 TaskService 置 FAILED。
     *
     * @param deviceId  前端设备号（X-Device-Id），多浏览器隔离
     */
    protected Map<String, Object> doSubmitAsync(MultipartFile file, String deviceId,
                                                String module, String apiName,
                                                String displayName, FileTaskWork work) {
        String originalName = file == null ? null : file.getOriginalFilename();
        try {
            byte[] bytes = file.getBytes();
            MemoryMultipartFile safe = new MemoryMultipartFile(
                    file.getName(), file.getOriginalFilename(), file.getContentType(), bytes);
            return taskService.submitAndRespond(deviceId, module, apiName, displayName,
                    originalName, () -> work.execute(safe));
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件读取失败：" + e.getMessage());
            return err;
        }
    }

    /** 取设备号（前端 X-Device-Id，缺失兼容单机/旧前端） */
    protected String deviceId(String header) {
        return (header == null || header.isBlank()) ? "" : header;
    }

    /** 把 MultipartFile 固化为内存副本（异步线程安全）：请求线程内调用，异步线程可安全 getBytes/getInputStream */
    protected MemoryMultipartFile pin(MultipartFile f) throws IOException {
        return new MemoryMultipartFile(f.getName(), f.getOriginalFilename(), f.getContentType(), f.getBytes());
    }

    protected ResponseEntity<Resource> doDownloadFile(String fileName) {
        try {
            File file = new File(uploadDir, fileName);
            if (!file.exists()) {
                return ResponseEntity.notFound().build();
            }
            Resource resource = new FileSystemResource(file);
            // RFC 5987：中文文件名需 UTF-8 百分号编码，否则 Tomcat 判定 Content-Disposition 非法直接移除该头
            String encoded = URLEncoder.encode(file.getName(), StandardCharsets.UTF_8).replace("+", "%20");
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=\"" + encoded + "\"; filename*=UTF-8''" + encoded)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(resource);
        } catch (Exception e) {
            return ResponseEntity.internalServerError().build();
        }
    }
}
