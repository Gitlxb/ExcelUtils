package com.example.project.controller.finance.jxkh;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.jxkh.JxkhSjppService;
import com.example.project.task.MemoryMultipartFile;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 绩效考核-数据匹配 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class JxkhSjppController extends FinanceBaseController {

    @Autowired
    private JxkhSjppService jxkhSjppService;

    /**
     * POST /api/finance/jxkh/sjpp
     * 功能：数据匹配（cw_sjpp_zhf.py）
     * 异步化（sourceFile 必传 + prevFile 可选）：提交即返回 taskId，后台线程池执行
     */
    @PostMapping("/jxkh/sjpp")
    public Map<String, Object> jxkhSjpp(
            @RequestParam("sourceFile") MultipartFile sourceFile,
            @RequestParam(value = "prevFile", required = false) MultipartFile prevFile,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            MemoryMultipartFile source = pin(sourceFile);
            MemoryMultipartFile prev = (prevFile == null || prevFile.isEmpty()) ? null : pin(prevFile);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "jxkh/sjpp", "绩效数据匹配",
                    sourceFile.getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        String outputPath = jxkhSjppService.generate(source, prev, uploadDir);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "数据匹配完成");
                        result.put("fileName", new File(outputPath).getName());
                        return result;
                    });
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件读取失败：" + e.getMessage());
            return err;
        }
    }
}
