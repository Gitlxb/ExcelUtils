package com.example.project.controller.hr;

import com.example.project.service.hr.CfDuogeWenjZzService;
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
public class CfDuogeWenjZzController extends HrBaseController {

    @Autowired
    private CfDuogeWenjZzService cfDuogeWenjZzService;

    /**
     * POST /api/hr/split-to-files
     * 功能：拆分 Excel 为多个独立文件（cf_duoge_wenj_zz.py）
     * 特点：长数字文本保护、日期格式化、字母倒序排序
     */
    @PostMapping("/split-to-files")
    public Map<String, Object> splitToFiles(
            @RequestParam("file") MultipartFile file,
            @RequestParam("columns") String columns) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = cfDuogeWenjZzService.splitToFiles(file, columns, uploadDir);
            result.put("success", true);
            result.put("message", "拆分文件完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
