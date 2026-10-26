package com.example.project.task;

import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

/**
 * 带安全文件副本的异步工作单元：BaseController.doSubmitAsync 在请求线程内
 * 将 MultipartFile 固化为 MemoryMultipartFile（内存副本，异步安全）后传入本回调。
 *
 * 单文件型工具（绝大多数）用它包装原 Service 调用；多文件/特殊工具直接
 * 用 MemoryMultipartFile 自行固化多个文件后调 taskService.submitAndRespond。
 */
@FunctionalInterface
public interface FileTaskWork {

    Map<String, Object> execute(MultipartFile safeFile) throws Exception;
}
