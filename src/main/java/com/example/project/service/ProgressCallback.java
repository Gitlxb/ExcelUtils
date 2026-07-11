package com.example.project.service;

/**
 * 进度回调接口，用于报告 Excel 处理进度
 */
@FunctionalInterface
public interface ProgressCallback {
    /**
     * 更新处理进度
     *
     * @param progress 进度百分比 (0-100)
     * @param message  当前处理状态描述
     */
    void updateProgress(int progress, String message);
}
