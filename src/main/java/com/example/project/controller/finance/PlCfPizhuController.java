package com.example.project.controller.finance;

import com.example.project.service.finance.PlCfPizhuService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 批量批注 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class PlCfPizhuController extends FinanceBaseController {

    @Autowired
    private PlCfPizhuService plCfPizhuService;

    /**
     * POST /api/finance/comment/add
     * 功能：批量添加批注（pl_cf_pizhu.py）
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/comment/add")
    public Map<String, Object> commentAdd(
            @RequestParam("file") MultipartFile file,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "comment/add", "批量批注添加",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = plCfPizhuService.addComments(safeFile, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "批注添加完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
    @PostMapping("/comment/split")
    public Map<String, Object> commentSplit(
            @RequestParam("file") MultipartFile file,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "comment/split", "批量批注拆分",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = plCfPizhuService.splitComments(safeFile, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "批注拆分完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
