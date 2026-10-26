package com.example.project.task;

import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.annotation.PostConstruct;
import jakarta.annotation.PreDestroy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Excel 工具「任务中心」核心服务（方案 A）。
 *
 * 职责：
 * 1. 接收异步任务：提交即落盘 PENDING → 固定线程池(2)后台执行 → 完成/失败落盘；
 * 2. 状态存取：内存 ConcurrentHashMap 热表 + uploads/.tasks/{taskId}.json 双写，
 *    启动时扫描 .tasks 恢复（后端重启不丢已完成记录）；
 * 3. 设备隔离：get/list 按 deviceId 过滤，非本人返回 null/过滤；
 * 4. 自动清理：任务结束(finishTime) 10 分钟后删除 json + 结果文件（@Scheduled 30s 扫）。
 *
 * 设计要点：
 * - 固定 2 线程 → 并发处理封顶，其余排队，JVM 内存峰值可控；
 * - 请求线程提交后立即返回（不 hold），POI 处理在 worker 线程完成，完成后线程归还；
 * - TaskWork 返回完整 result Map（含 fileName 等），与原同步 res.data 同构。
 */
@Service
public class TaskService {

    private static final Logger log = LoggerFactory.getLogger(TaskService.class);

    public static final String STATUS_PENDING = "PENDING";
    public static final String STATUS_RUNNING = "RUNNING";
    public static final String STATUS_SUCCESS = "SUCCESS";
    public static final String STATUS_FAILED = "FAILED";

    /** 任务结束后保留时长：10 分钟 */
    private static final long KEEP_AFTER_FINISH_MS = 10L * 60 * 1000;
    /** 并发 worker 数：内存封顶，避免多个大 Excel 叠加 OOM */
    private static final int WORKER_COUNT = 2;
    /** PENDING 排队超过 10 分钟未轮到（极端情况）也清理 */
    private static final long PENDING_MAX_WAIT_MS = 10L * 60 * 1000;

    private final Map<String, TaskRecord> tasks = new ConcurrentHashMap<>();
    private final ObjectMapper objectMapper = new ObjectMapper();
    private final ExecutorService executor;

    @Value("${file.upload-dir:./uploads}")
    private String uploadDir;

    public TaskService() {
        ThreadFactory tf = new ThreadFactory() {
            private final AtomicInteger seq = new AtomicInteger(1);
            @Override
            public Thread newThread(Runnable r) {
                Thread t = new Thread(r, "task-worker-" + seq.getAndIncrement());
                t.setDaemon(true);
                return t;
            }
        };
        this.executor = Executors.newFixedThreadPool(WORKER_COUNT, tf);
    }

    // ==================== 提交与执行 ====================

    /**
     * 提交异步任务，立即返回 taskId。
     *
     * @param deviceId     设备号（可 null，视为单机兼容）
     * @param module       general | finance | hr
     * @param apiName      接口名（split-to-sheets 等）
     * @param displayName  工具中文名
     * @param params       参数 JSON（可 null）
     * @param originalFileName 上传源文件名（可 null）
     * @param work         实际处理逻辑（返回完整 result Map）
     */
    public String submit(String deviceId, String module, String apiName, String displayName,
                         String params, String originalFileName, TaskWork work) {
        String taskId = UUID.randomUUID().toString().replace("-", "");
        TaskRecord r = new TaskRecord();
        r.setTaskId(taskId);
        r.setDeviceId(deviceId == null ? "" : deviceId);
        r.setModule(module);
        r.setApiName(apiName);
        r.setDisplayName(displayName);
        r.setParams(params);
        r.setOriginalFileName(originalFileName);
        r.setStatus(STATUS_PENDING);
        r.setCreateTime(System.currentTimeMillis());
        tasks.put(taskId, r);
        saveToDisk(r);
        log.info("[任务中心] 提交 taskId={} module={} api={} display={} device={}",
                taskId, module, apiName, displayName, r.getDeviceId());

        executor.submit(() -> run(taskId, work));
        return taskId;
    }

    /** Controller 便捷方法：提交并返回 {success,taskId}（提交异常时返回 {success:false,message}） */
    public Map<String, Object> submitAndRespond(String deviceId, String module, String apiName,
                                                String displayName, String originalFileName,
                                                TaskWork work) {
        Map<String, Object> res = new HashMap<>();
        try {
            String taskId = submit(deviceId, module, apiName, displayName, null,
                    originalFileName, work);
            res.put("success", true);
            res.put("taskId", taskId);
            res.put("message", "任务已提交，后台处理中");
        } catch (Exception e) {
            res.put("success", false);
            res.put("message", "提交失败：" + e.getMessage());
        }
        return res;
    }

    private void run(String taskId, TaskWork work) {
        TaskRecord r = tasks.get(taskId);
        if (r == null) return;
        r.setStatus(STATUS_RUNNING);
        r.setMessage(null);
        saveToDisk(r);
        log.info("[任务中心] 开始执行 taskId={}", taskId);
        try {
            Map<String, Object> result = work.execute();
            if (result == null) {
                result = new HashMap<>();
            }
            result.putIfAbsent("success", true);
            r.setResult(result);
            // 业务方可在 result 显式放 success=false 并附带 logs/warnings → 状态置 FAILED 且保留结果体
            if (Boolean.FALSE.equals(result.get("success"))) {
                r.setStatus(STATUS_FAILED);
                r.setMessage(result.get("message") == null ? "处理未完成" : String.valueOf(result.get("message")));
            } else {
                r.setStatus(STATUS_SUCCESS);
                Object fn = result.get("fileName");
                log.info("[任务中心] 完成 taskId={} status=SUCCESS fileName={}", taskId, fn);
            }
        } catch (Throwable t) {
            // 捕获所有异常（含 Error），绝不丢任务状态
            r.setStatus(STATUS_FAILED);
            String msg = t.getMessage();
            if (msg == null || msg.isBlank()) {
                msg = t.getClass().getSimpleName();
            }
            r.setMessage("处理失败：" + msg);
            log.error("[任务中心] 失败 taskId={}", taskId, t);
        } finally {
            r.setFinishTime(System.currentTimeMillis());
            saveToDisk(r);
        }
    }

    // ==================== 查询 ====================

    /** 查询单个任务；deviceId 不匹配返回 null（前端视为无权访问） */
    public TaskRecord get(String taskId, String deviceId) {
        TaskRecord r = tasks.get(taskId);
        if (r == null) return null;
        String d = deviceId == null ? "" : deviceId;
        if (!r.getDeviceId().equals(d)) return null;
        return r;
    }

    /** 列出某设备全部未清理任务（PENDING/RUNNING + 未过期已完成），按创建时间倒序 */
    public List<TaskRecord> listByDevice(String deviceId) {
        String d = deviceId == null ? "" : deviceId;
        long now = System.currentTimeMillis();
        List<TaskRecord> list = new ArrayList<>();
        for (TaskRecord r : tasks.values()) {
            if (!r.getDeviceId().equals(d)) continue;
            if (isExpired(r, now)) continue;   // 过期未清的不再展示
            list.add(r);
        }
        list.sort(Comparator.comparingLong(TaskRecord::getCreateTime).reversed());
        return list;
    }

    private boolean isExpired(TaskRecord r, long now) {
        if (STATUS_RUNNING.equals(r.getStatus())) return false;       // 运行中绝不清
        if (STATUS_PENDING.equals(r.getStatus())) {
            return now - r.getCreateTime() > PENDING_MAX_WAIT_MS;     // 排队超时视为过期
        }
        // SUCCESS / FAILED：从结束时刻起 10 分钟
        return r.getFinishTime() > 0 && now - r.getFinishTime() > KEEP_AFTER_FINISH_MS;
    }

    // ==================== 持久化 ====================

    private Path tasksDir() {
        return Path.of(uploadDir, ".tasks");
    }

    private void saveToDisk(TaskRecord r) {
        try {
            Files.createDirectories(tasksDir());
            Path p = tasksDir().resolve(r.getTaskId() + ".json");
            objectMapper.writeValue(p.toFile(), r);
        } catch (IOException e) {
            log.warn("[任务中心] 任务落盘失败 taskId={}", r.getTaskId(), e);
        }
    }

    /** 启动恢复：扫描 .tasks/*.json 载入内存表（后端重启后已完成记录仍可查/下载） */
    @PostConstruct
    public void loadFromDisk() {
        Path dir = tasksDir();
        if (!Files.isDirectory(dir)) return;
        File[] files = dir.toFile().listFiles((d, n) -> n.endsWith(".json"));
        if (files == null) return;
        int loaded = 0;
        for (File f : files) {
            try {
                TaskRecord r = objectMapper.readValue(f, TaskRecord.class);
                if (r != null && r.getTaskId() != null) {
                    // 运行中的任务（进程重启导致中断）标记失败，避免永远 RUNNING
                    if (STATUS_RUNNING.equals(r.getStatus()) || STATUS_PENDING.equals(r.getStatus())) {
                        r.setStatus(STATUS_FAILED);
                        r.setMessage("服务重启，任务中断（原状态 " + r.getStatus() + "）");
                        r.setFinishTime(System.currentTimeMillis());
                        saveToDisk(r);
                    }
                    tasks.put(r.getTaskId(), r);
                    loaded++;
                }
            } catch (IOException e) {
                log.warn("[任务中心] 任务 json 解析失败，跳过 {}", f.getName(), e);
            }
        }
        log.info("[任务中心] 启动恢复 {} 条任务记录", loaded);
    }

    // ==================== 清理 ====================

    /** 每 30s 扫描：清理结束超过 10 分钟的任务（json + 结果文件） */
    @Scheduled(fixedDelay = 30_000)
    public void cleanup() {
        long now = System.currentTimeMillis();
        int removed = 0;
        List<String> expiredIds = new ArrayList<>();
        for (TaskRecord r : tasks.values()) {
            if (isExpired(r, now)) {
                expiredIds.add(r.getTaskId());
                deleteResultFile(r);
            }
        }
        for (String id : expiredIds) {
            tasks.remove(id);
            try {
                Files.deleteIfExists(tasksDir().resolve(id + ".json"));
            } catch (IOException e) {
                log.warn("[任务中心] 清理 json 失败 taskId={}", id, e);
            }
            removed++;
        }
        if (removed > 0) {
            log.info("[任务中心] 清理过期任务 {} 条（保留窗口 10 分钟）", removed);
        }
    }

    private void deleteResultFile(TaskRecord r) {
        if (r.getResult() == null) return;
        Object fn = r.getResult().get("fileName");
        if (fn == null || fn.toString().isBlank()) return;
        String name = fn.toString();
        // 防目录穿越：仅允许单一文件名
        if (name.contains("/") || name.contains("\\") || name.contains("..")) return;
        try {
            Files.deleteIfExists(Path.of(uploadDir).resolve(name));
        } catch (IOException e) {
            log.warn("[任务中心] 清理结果文件失败 name={}", name, e);
        }
    }

    @PreDestroy
    public void shutdown() {
        executor.shutdownNow();
        log.info("[任务中心] 线程池已关闭");
    }
}
