package com.example.project.controller.general;

import com.example.project.service.general.CfBiaogeLiService;
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
public class CfBiaogeLiController extends GeneralBaseController {

    @Autowired
    private CfBiaogeLiService cfBiaogeLiService;

    /**
     * POST /api/general/split-to-sheets
     * 功能：按列拆分为多个工作表（cf_biaoge_li.py）
     * 异步化：提交即返回 taskId，后台线程池执行；前端轮询 /api/tasks/{taskId} 拿结果。
     */
    @PostMapping("/split-to-sheets")
    public Map<String, Object> splitToSheets(
            @RequestParam("file") MultipartFile file,
            @RequestParam("column") String column,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "general", "split-to-sheets", "拆成多个工作表",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = cfBiaogeLiService.splitToSheets(safeFile, column, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "拆分工作表完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }
}
