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
     */
    @PostMapping("/jxkh/wgl")
    public Map<String, Object> jxkhWgl(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "sheetName", required = false) String sheetName,
            @RequestParam("startMonth") String startMonth,
            @RequestParam("endMonth") String endMonth,
            @RequestParam(value = "copyOriginal", defaultValue = "true") boolean copyOriginal) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            JxkhWglService.AnalyzeResult ar = jxkhWglService.analyze(
                    file, sheetName, startMonth, endMonth, copyOriginal, uploadDir);
            result.put("success", true);
            result.put("message", "稳岗率分析完成");
            result.put("fileName", new File(ar.filePath).getName());
            result.put("logs", ar.logs);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
