package com.example.project.task;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.Map;

/**
 * 异步任务记录（内存热表 + uploads/.tasks/{taskId}.json 落盘，两者同构）。
 *
 * 字段即对外返回体（TaskController 直接序列化本对象），前端据此渲染状态与结果。
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public class TaskRecord {

    /** 任务ID（UUID 无横线） */
    private String taskId;

    /** 归属设备号（前端 X-Device-Id 传入），列表/查询按此隔离 */
    private String deviceId;

    /** 模块：general | finance | hr（拼下载 URL 用） */
    private String module;

    /** 接口名：split-to-sheets 等 */
    private String apiName;

    /** 工具中文名（任务中心展示） */
    private String displayName;

    /** 原始参数 JSON（列名等，痕迹展示用） */
    private String params;

    /** 上传的源文件名 */
    private String originalFileName;

    /** PENDING | RUNNING | SUCCESS | FAILED */
    private String status;

    /** 成功/失败文案 */
    private String message;

    /**
     * 成功后的完整返回体（与原同步 res.data 同构）：
     * 至少含 fileName；特殊工具含 summaryFileNames / splitCount / step4Details 等自定义字段。
     */
    private Map<String, Object> result;

    /** 创建时间 epoch ms */
    private long createTime;

    /** 结束时间 epoch ms（清理从此时起算 10 分钟） */
    private long finishTime;

    public String getTaskId() { return taskId; }
    public void setTaskId(String taskId) { this.taskId = taskId; }

    public String getDeviceId() { return deviceId; }
    public void setDeviceId(String deviceId) { this.deviceId = deviceId; }

    public String getModule() { return module; }
    public void setModule(String module) { this.module = module; }

    public String getApiName() { return apiName; }
    public void setApiName(String apiName) { this.apiName = apiName; }

    public String getDisplayName() { return displayName; }
    public void setDisplayName(String displayName) { this.displayName = displayName; }

    public String getParams() { return params; }
    public void setParams(String params) { this.params = params; }

    public String getOriginalFileName() { return originalFileName; }
    public void setOriginalFileName(String originalFileName) { this.originalFileName = originalFileName; }

    public String getStatus() { return status; }
    public void setStatus(String status) { this.status = status; }

    public String getMessage() { return message; }
    public void setMessage(String message) { this.message = message; }

    public Map<String, Object> getResult() { return result; }
    public void setResult(Map<String, Object> result) { this.result = result; }

    public long getCreateTime() { return createTime; }
    public void setCreateTime(long createTime) { this.createTime = createTime; }

    public long getFinishTime() { return finishTime; }
    public void setFinishTime(long finishTime) { this.finishTime = finishTime; }
}
