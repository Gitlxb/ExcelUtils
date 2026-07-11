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
     */
    @PostMapping("/comment/add")
    public Map<String, Object> commentAdd(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = plCfPizhuService.addComments(file, uploadDir);
            result.put("success", true);
            result.put("message", "批注添加完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/comment/split
     * 功能：批量拆分批注（pl_cf_pizhu.py）
     */
    @PostMapping("/comment/split")
    public Map<String, Object> commentSplit(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = plCfPizhuService.splitComments(file, uploadDir);
            result.put("success", true);
            result.put("message", "批注拆分完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
