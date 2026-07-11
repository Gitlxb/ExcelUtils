package com.example.project.controller.finance;

import com.example.project.service.finance.CwYhlsZyxgService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 银行流水处理 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwYhlsZyxgController extends FinanceBaseController {

    @Autowired
    private CwYhlsZyxgService cwYhlsZyxgService;

    /**
     * POST /api/finance/bank-flow/detect
     * 功能：银行流水-检测列配置
     */
    @PostMapping("/bank-flow/detect")
    public Map<String, Object> bankFlowDetect(@RequestParam("file") MultipartFile file) {
        try {
            return cwYhlsZyxgService.detectConfig(file);
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "检测失败：" + e.getMessage());
            return err;
        }
    }

    /**
     * POST /api/finance/bank-flow/process
     * 功能：银行流水-执行处理（cw_yhls_zyxg.py）
     */
    @PostMapping("/bank-flow/process")
    public Map<String, Object> bankFlowProcess(
            @RequestParam("file") MultipartFile file,
            @RequestParam("selectedAs") String selectedAs,
            @RequestParam("selectedCompany") String selectedCompany,
            @RequestParam(value = "inputNameFilter", required = false) String inputNameFilter,
            @RequestParam(value = "inputSummaryText", required = false) String inputSummaryText,
            @RequestParam(value = "inputSummary", required = false) String inputSummary) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwYhlsZyxgService.process(
                    file, selectedAs, selectedCompany,
                    inputNameFilter, inputSummaryText, inputSummary, uploadDir);
            result.put("success", true);
            result.put("message", "银行流水处理完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
