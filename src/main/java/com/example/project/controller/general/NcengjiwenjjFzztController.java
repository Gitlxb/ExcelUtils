package com.example.project.controller.general;

import com.example.project.service.general.NcengjiwenjjFzztService;
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
public class NcengjiwenjjFzztController extends GeneralBaseController {

    @Autowired
    private NcengjiwenjjFzztService ncengjiwenjjFzztService;

    /**
     * POST /api/general/extract-files
     * 功能：文件夹文件提取（异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}）
     */
    @PostMapping("/extract-files")
    public Map<String, Object> extractFiles(
            @RequestParam("file") MultipartFile file,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "general", "extract-files", "文件夹文件提取",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = ncengjiwenjjFzztService.extractFilesFromZip(safeFile, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "文件提取完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
