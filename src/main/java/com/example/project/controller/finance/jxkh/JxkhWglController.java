package com.example.project.controller.finance.jxkh;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.jxkh.JxkhWglService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 绩效考核-稳岗率分析 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class JxkhWglController extends FinanceBaseController {

    @Autowired
    private JxkhWglService jxkhWglService;

    /**
     * POST /api/finance/jxkh/wgl
     * 功能：稳岗率分析（cw_wgl_zhf.py）
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/jxkh/wgl")
    public Map<String, Object> jxkhWgl(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "sheetName", required = false) String sheetName,
            @RequestParam("startMonth") String startMonth,
            @RequestParam("endMonth") String endMonth,
            @RequestParam(value = "copyOriginal", defaultValue = "true") boolean copyOriginal,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "jxkh/wgl", "稳岗率分析",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    JxkhWglService.AnalyzeResult ar = jxkhWglService.analyze(
                            safeFile, sheetName, startMonth, endMonth, copyOriginal, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "稳岗率分析完成");
                    result.put("fileName", new File(ar.filePath).getName());
                    result.put("logs", ar.logs);
                    return result;
                });
    }
}
