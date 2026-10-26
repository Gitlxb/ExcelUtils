package com.example.project.controller.finance.cwgjx;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cwgjx.CwgjxBijiaoService;
import com.example.project.task.MemoryMultipartFile;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 数据比较 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwgjxBijiaoController extends FinanceBaseController {

    @Autowired
    private CwgjxBijiaoService cwgjxBijiaoService;

    /**
     * POST /api/finance/toolbox/data-compare
     * 功能：数据比较（cw_bijiao_xzy.py）
     * 异步化（双文件）：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/toolbox/data-compare")
    public Map<String, Object> toolboxDataCompare(
            @RequestParam("file1") MultipartFile file1,
            @RequestParam("file2") MultipartFile file2,
            @RequestParam("sheet1Name") String sheet1Name,
            @RequestParam("sheet2Name") String sheet2Name,
            @RequestParam("startDate") String startDate,
            @RequestParam("endDate") String endDate,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            MemoryMultipartFile f1 = pin(file1);
            MemoryMultipartFile f2 = pin(file2);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "toolbox/data-compare", "数据比较",
                    file1.getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        CwgjxBijiaoService.CompareResult compareResult = cwgjxBijiaoService.compare(
                                f1, sheet1Name, f2, sheet2Name, startDate, endDate, uploadDir);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "数据比较完成");
                        result.put("fileName", compareResult.fileName);
                        result.put("preview", compareResult.preview);
                        result.put("previewRows", compareResult.previewRows);
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
