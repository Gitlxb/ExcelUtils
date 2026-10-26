package com.example.project.controller.finance;

import com.example.project.service.finance.CwKgzSxService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 开关账筛选 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwKgzSxController extends FinanceBaseController {

    @Autowired
    private CwKgzSxService cwKgzSxService;

    /**
     * POST /api/finance/date-filter
     * 功能：按日期范围筛选（cw_kgz_sx.py）
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/date-filter")
    public Map<String, Object> dateFilter(
            @RequestParam("file") MultipartFile file,
            @RequestParam("dateColumn") String dateColumn,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "date-filter", "开关账筛选",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = cwKgzSxService.filterByDate(safeFile, dateColumn, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "日期筛选完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
