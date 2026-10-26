package com.example.project.task;

import java.util.Map;

/**
 * 异步任务工作单元：包装原 Service 的同步处理调用。
 *
 * 约定：返回的 Map 即原同步 Controller 成功路径的 result（含 fileName；
 * 特殊工具如 labor-split-batch 可含 summaryFileNames/splitCount/step4Details 等），
 * TaskService 会将其存入 TaskRecord.result，前端轮询到 SUCCESS 后取用，
 * 字段结构与改造前 res.data 完全一致 → 页面渲染零改动。
 *
 * 异常处理：方法抛出的任何异常由 TaskService 捕获，任务置 FAILED 并记录 message。
 */
@FunctionalInterface
public interface TaskWork {

    Map<String, Object> execute() throws Exception;
}
