package com.example.project.controller.general;

import com.example.project.service.general.ZhPdfExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
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
     * 功能：PDF转Excel（异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}）
     */
    @PostMapping("/pdf-to-excel")
    public Map<String, Object> pdfToExcel(
            @RequestParam("file") MultipartFile file,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "general", "pdf-to-excel", "PDF转Excel",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = zhPdfExcelService.convertPdfToExcel(safeFile, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "PDF 转换完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
