package com.example.project.task;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 任务中心查询接口。
 *
 * 设备号经请求头 X-Device-Id 传入（前端 axios 拦截器统一附加）；
 * 缺失时按空设备处理（单机/未升级前端的兼容）。
 * 下载仍走原 /api/{module}/download/{fileName}，不在此重复。
 */
@RestController
@RequestMapping("/api/tasks")
public class TaskController {

    private final TaskService taskService;

    public TaskController(TaskService taskService) {
        this.taskService = taskService;
    }

    /** 查询单任务状态：GET /api/tasks/{taskId}（带 X-Device-Id） */
    @GetMapping("/{taskId}")
    public Map<String, Object> get(@PathVariable String taskId,
                                   @RequestHeader(value = "X-Device-Id", required = false) String deviceId) {
        TaskRecord r = taskService.get(taskId, deviceId);
        Map<String, Object> res = new HashMap<>();
        if (r == null) {
            res.put("success", false);
            res.put("message", "任务不存在或无权访问");
            return res;
        }
        res.put("success", true);
        res.put("task", r);
        return res;
    }

    /** 列出本设备任务：GET /api/tasks（带 X-Device-Id） */
    @GetMapping
    public Map<String, Object> list(@RequestHeader(value = "X-Device-Id", required = false) String deviceId) {
        List<TaskRecord> list = taskService.listByDevice(deviceId);
        Map<String, Object> res = new HashMap<>();
        res.put("success", true);
        res.put("tasks", list);
        return res;
    }
}
