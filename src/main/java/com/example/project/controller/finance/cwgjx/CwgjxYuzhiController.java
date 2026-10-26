package com.example.project.controller.finance.cwgjx;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cwgjx.CwgjxYuzhiService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 预支处理 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwgjxYuzhiController extends FinanceBaseController {

    @Autowired
    private CwgjxYuzhiService cwgjxYuzhiService;

    /**
     * POST /api/finance/toolbox/fee-split
     * 功能：预支处理（yuzhi_frame.py）
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/toolbox/fee-split")
    public Map<String, Object> toolboxFeeSplit(@RequestParam("file") MultipartFile file,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "toolbox/fee-split", "预支处理",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = cwgjxYuzhiService.process(safeFile, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "预支处理完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
