package com.example.project.controller.finance;

import com.example.project.service.finance.CwGysQzyService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 供应商合并 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwGysQzyController extends FinanceBaseController {

    @Autowired
    private CwGysQzyService cwGysQzyService;

    /**
     * POST /api/finance/merge-supplier
     * 功能：供应商 Excel 合并（cw_gys_qzy.py）
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/merge-supplier")
    public Map<String, Object> mergeSupplier(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "merge-supplier", "供应商合并",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = cwGysQzyService.mergeExcelFiles(safeFile, sheetName, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "合并完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
