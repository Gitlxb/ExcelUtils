package com.example.project.controller.finance;

import com.example.project.service.finance.CwGysQzyService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 供应商合并 Controller
 */
@RestController
@RequestMapping("/api/finance")
public class CwGysQzyController extends FinanceBaseController {

    @Autowired
    private CwGysQzyService cwGysQzyService;

    /**
     * POST /api/finance/merge-supplier
     * 功能：供应商 Excel 合并（cw_gys_qzy.py）
     */
    @PostMapping("/merge-supplier")
    public Map<String, Object> mergeSupplier(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = cwGysQzyService.mergeExcelFiles(file, sheetName, uploadDir);
            result.put("success", true);
            result.put("message", "合并完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
