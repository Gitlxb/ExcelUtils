package com.example.project.controller.finance.cwgjx;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cwgjx.CwgjxBxcfService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 报销拆分 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwgjxBxcfController extends FinanceBaseController {

    @Autowired
    private CwgjxBxcfService cwgjxBxcfService;

    /**
     * POST /api/finance/toolbox/expense-preview
     * 功能：报销拆分-预览列名（bxcf_frame.py 前三步）
     */
    @PostMapping("/toolbox/expense-preview")
    public Map<String, Object> toolboxExpensePreview(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            List<String> columns = cwgjxBxcfService.processPreview(file, uploadDir);
            result.put("success", true);
            result.put("message", "预览处理完成");
            result.put("columns", columns);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/toolbox/expense-split
     * 功能：报销拆分-执行拆分（bxcf_frame.py）
     */
    @PostMapping("/toolbox/expense-split")
    public Map<String, Object> toolboxExpenseSplit(
            @RequestParam("file") MultipartFile file,
            @RequestParam("splitColumn") String splitColumn) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwgjxBxcfService.process(file, splitColumn, uploadDir);
            result.put("success", true);
            result.put("message", "报销拆分完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
