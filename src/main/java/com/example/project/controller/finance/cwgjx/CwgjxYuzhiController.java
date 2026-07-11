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
 * 手续费拆分 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwgjxYuzhiController extends FinanceBaseController {

    @Autowired
    private CwgjxYuzhiService cwgjxYuzhiService;

    /**
     * POST /api/finance/toolbox/fee-split
     * 功能：手续费拆分（yuzhi_frame.py）
     */
    @PostMapping("/toolbox/fee-split")
    public Map<String, Object> toolboxFeeSplit(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwgjxYuzhiService.process(file, uploadDir);
            result.put("success", true);
            result.put("message", "手续费拆分完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
