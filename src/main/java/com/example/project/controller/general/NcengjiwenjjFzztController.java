package com.example.project.controller.general;

import com.example.project.service.general.NcengjiwenjjFzztService;
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
public class NcengjiwenjjFzztController extends GeneralBaseController {

    @Autowired
    private NcengjiwenjjFzztService ncengjiwenjjFzztService;

    /**
     * POST /api/general/extract-files
     * 功能：从ZIP中提取嵌套文件夹的所有文件（Ncengjiwenjj_fzzt.py）
     */
    @PostMapping("/extract-files")
    public Map<String, Object> extractFiles(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = ncengjiwenjjFzztService.extractFilesFromZip(file, uploadDir);
            result.put("success", true);
            result.put("message", "文件提取完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
