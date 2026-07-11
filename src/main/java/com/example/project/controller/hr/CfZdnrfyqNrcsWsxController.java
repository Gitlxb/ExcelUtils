package com.example.project.controller.hr;

import com.example.project.service.hr.CfZdnrfyqNrcsWsxService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/api/hr")
public class CfZdnrfyqNrcsWsxController extends HrBaseController {

    @Autowired
    private CfZdnrfyqNrcsWsxService cfZdnrfyqNrcsWsxService;

    /**
     * POST /api/hr/split-by-config
     * 功能：按"拆分参数"工作表配置自动拆分（cf_zdnrfyq_nrcs_wsx.py）
     */
    @PostMapping("/split-by-config")
    public Map<String, Object> splitByConfig(
            @RequestParam("file") MultipartFile file,
            @RequestParam("column") String column) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = cfZdnrfyqNrcsWsxService.splitByConfig(file, column, uploadDir);
            result.put("success", true);
            result.put("message", "按配置拆分完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
