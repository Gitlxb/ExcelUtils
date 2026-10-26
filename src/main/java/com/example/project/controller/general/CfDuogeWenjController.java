package com.example.project.controller.general;

import com.example.project.service.general.CfDuogeWenjService;
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
public class CfDuogeWenjController extends GeneralBaseController {

    @Autowired
    private CfDuogeWenjService cfDuogeWenjService;

    /**
     * POST /api/general/split-to-files
     * 功能：拆成多个Excel文件（异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}）
     */
    @PostMapping("/split-to-files")
    public Map<String, Object> splitToFiles(
            @RequestParam("file") MultipartFile file,
            @RequestParam("columns") String columns,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "general", "split-to-files", "拆成多个Excel文件",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = cfDuogeWenjService.splitToFiles(safeFile, columns, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "拆分文件完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
