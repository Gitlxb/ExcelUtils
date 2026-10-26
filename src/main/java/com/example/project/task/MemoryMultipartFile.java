package com.example.project.task;

import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * 内存版 MultipartFile：在请求线程内用 file.getBytes() 固化后交给异步任务线程使用。
 *
 * 背景：Spring 容器管理的 MultipartFile（StandardMultipartFile）生命周期绑定当前请求，
 * 请求线程返回后底层 Part/临时文件可能被清理，直接传入 TaskService 线程池存在数据失效风险。
 * 本类把内容持在 byte[]，getInputStream/transferTo 均基于内存，不依赖请求上下文，异步安全。
 *
 * 用法（Controller 包装模板）：
 *   byte[] bytes = file.getBytes();
 *   MultipartFile safe = new MemoryMultipartFile("file", file.getOriginalFilename(),
 *                                                 file.getContentType(), bytes);
 *   taskService.submit(..., () -> service.xxx(safe, params, uploadDir));
 */
public class MemoryMultipartFile implements MultipartFile {

    private final String name;
    private final String originalFilename;
    private final String contentType;
    private final byte[] content;

    public MemoryMultipartFile(String name, String originalFilename, String contentType, byte[] content) {
        this.name = name;
        this.originalFilename = originalFilename;
        this.contentType = contentType;
        this.content = content == null ? new byte[0] : content;
    }

    @Override
    public String getName() { return name; }

    @Override
    public String getOriginalFilename() { return originalFilename; }

    @Override
    public String getContentType() { return contentType; }

    @Override
    public boolean isEmpty() { return content.length == 0; }

    @Override
    public long getSize() { return content.length; }

    @Override
    public byte[] getBytes() throws IOException { return content; }

    @Override
    public InputStream getInputStream() throws IOException {
        return new ByteArrayInputStream(content);
    }

    @Override
    public void transferTo(File dest) throws IOException, IllegalStateException {
        java.nio.file.Files.write(dest.toPath(), content);
    }
}
