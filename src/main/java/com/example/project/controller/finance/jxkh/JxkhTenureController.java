package com.example.project.controller.finance.jxkh;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.jxkh.JxkhTenureService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 绩效考核 - 员工在职月份计算 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class JxkhTenureController extends FinanceBaseController {

    @Autowired
    private JxkhTenureService jxkhTenureService;

    /**
     * POST /api/finance/jxkh/tenure
     * 功能：员工在职月份计算（tenure_months_calculator.py）
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/jxkh/tenure")
    public Map<String, Object> jxkhTenure(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "sheetName", required = false) String sheetName,
            @RequestParam(value = "startDate", required = false) String startDate,
            @RequestParam(value = "endDate", required = false) String endDate,
            @RequestParam(value = "mergeTotalDays", defaultValue = "false") boolean mergeTotalDays,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "jxkh/tenure", "在职月份计算",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    JxkhTenureService.AnalyzeResult ar = jxkhTenureService.analyze(
                            safeFile, sheetName, startDate, endDate, mergeTotalDays, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "在职月份计算完成");
                    result.put("fileName", new File(ar.filePath).getName());
                    result.put("logs", ar.logs);
                    return result;
                });
    }
}
