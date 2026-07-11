package com.example.project.controller.finance.jxkh;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.jxkh.JxkhSjppService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 绩效考核-数据匹配 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class JxkhSjppController extends FinanceBaseController {

    @Autowired
    private JxkhSjppService jxkhSjppService;

    /**
     * POST /api/finance/jxkh/sjpp
     * 功能：数据匹配（cw_sjpp_zhf.py）
     */
    @PostMapping("/jxkh/sjpp")
    public Map<String, Object> jxkhSjpp(
            @RequestParam("sourceFile") MultipartFile sourceFile,
            @RequestParam(value = "prevFile", required = false) MultipartFile prevFile) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = jxkhSjppService.generate(sourceFile, prevFile, uploadDir);
            result.put("success", true);
            result.put("message", "数据匹配完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
