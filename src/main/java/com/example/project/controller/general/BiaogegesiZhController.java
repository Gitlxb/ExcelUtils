package com.example.project.controller.general;

import com.example.project.service.general.BiaogegesiZhService;
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
public class BiaogegesiZhController extends GeneralBaseController {

    @Autowired
    private BiaogegesiZhService biaogegesiZhService;

    /**
     * POST /api/general/convert-format
     * 功能：表格格式转换（异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}）
     */
    @PostMapping("/convert-format")
    public Map<String, Object> convertFormat(
            @RequestParam("file") MultipartFile file,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "general", "convert-format", "表格格式转换",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = biaogegesiZhService.convertToXlsx(safeFile, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "格式转换完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
