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
     */
    @PostMapping("/jxkh/tenure")
    public Map<String, Object> jxkhTenure(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "sheetName", required = false) String sheetName,
            @RequestParam("statMonth") String statMonth,
            @RequestParam(value = "asOfDate", required = false) String asOfDate,
            @RequestParam(value = "mergeTotalDays", defaultValue = "true") boolean mergeTotalDays) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            JxkhTenureService.AnalyzeResult ar = jxkhTenureService.analyze(
                    file, sheetName, statMonth, asOfDate, mergeTotalDays, uploadDir);
            result.put("success", true);
            result.put("message", "在职月份计算完成");
            result.put("fileName", new File(ar.filePath).getName());
            result.put("logs", ar.logs);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
