package com.example.project.controller.general;

import com.example.project.service.general.BiaogegesiZhService;
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
@RequestMapping("/api/general")
public class BiaogegesiZhController extends GeneralBaseController {

    @Autowired
    private BiaogegesiZhService biaogegesiZhService;

    /**
     * POST /api/general/convert-format
     * 功能：表格格式转换（biaogegesi_zh.py）
     * 将 ZIP 包内的 .xls/.csv 等批量转为 .xlsx
     */
    @PostMapping("/convert-format")
    public Map<String, Object> convertFormat(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = biaogegesiZhService.convertToXlsx(file, uploadDir);
            result.put("success", true);
            result.put("message", "格式转换完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
