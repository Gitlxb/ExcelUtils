package com.example.project.controller.general;

import com.example.project.service.general.CfDuogeWenjService;
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
public class CfDuogeWenjController extends GeneralBaseController {

    @Autowired
    private CfDuogeWenjService cfDuogeWenjService;

    /**
     * POST /api/general/split-to-files
     * 功能：按列拆分为多个独立 Excel 文件，打包为 ZIP（cf_duoge_wenj.py）
     */
    @PostMapping("/split-to-files")
    public Map<String, Object> splitToFiles(
            @RequestParam("file") MultipartFile file,
            @RequestParam("columns") String columns) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = cfDuogeWenjService.splitToFiles(file, columns, uploadDir);
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
