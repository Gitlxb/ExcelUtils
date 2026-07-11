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
     */
    @PostMapping("/date-filter")
    public Map<String, Object> dateFilter(
            @RequestParam("file") MultipartFile file,
            @RequestParam("dateColumn") String dateColumn) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = cwKgzSxService.filterByDate(file, dateColumn, uploadDir);
            result.put("success", true);
            result.put("message", "日期筛选完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Throwable t) {
            result.put("success", false);
            String msg = t.getMessage();
            if (msg == null || msg.isEmpty()) msg = t.getClass().getSimpleName();
            result.put("message", "处理失败：" + msg);
            t.printStackTrace();
        }
        return result;
    }
}
