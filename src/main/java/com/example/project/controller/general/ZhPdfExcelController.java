package com.example.project.controller.general;

import com.example.project.service.general.ZhPdfExcelService;
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
public class ZhPdfExcelController extends GeneralBaseController {

    @Autowired
    private ZhPdfExcelService zhPdfExcelService;

    /**
     * POST /api/general/pdf-to-excel
     * 功能：PDF 转 Excel（zh_pdf_Excel.py）
     */
    @PostMapping("/pdf-to-excel")
    public Map<String, Object> pdfToExcel(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = zhPdfExcelService.convertPdfToExcel(file, uploadDir);
            result.put("success", true);
            result.put("message", "PDF 转换完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
