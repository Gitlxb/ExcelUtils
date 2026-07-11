package com.example.project.controller.finance;

import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

/**
 * Finance 模块公用端点 Controller
 * 将 BaseController 中的共享端点集中注册，避免被子 Controller 继承导致重复映射
 */
@RestController
@RequestMapping("/api/finance")
public class FinanceUtilController extends FinanceBaseController {

    @GetMapping("/download/{fileName}")
    public ResponseEntity<Resource> downloadFile(@PathVariable String fileName) {
        return doDownloadFile(fileName);
    }

    @PostMapping("/toolbox/sheets")
    public Map<String, Object> toolboxSheets(@RequestParam("file") MultipartFile file) {
        return doToolboxSheets(file);
    }

    @PostMapping("/summary/sheets")
    public Map<String, Object> summarySheets(@RequestParam("file") MultipartFile file) {
        return doSummarySheets(file);
    }
}
