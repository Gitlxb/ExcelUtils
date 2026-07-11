package com.example.project.controller.finance.cwgjx;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cwgjx.CwgjxBijiaoService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
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
     */
    @PostMapping("/toolbox/data-compare")
    public Map<String, Object> toolboxDataCompare(
            @RequestParam("file1") MultipartFile file1,
            @RequestParam("file2") MultipartFile file2,
            @RequestParam("sheet1Name") String sheet1Name,
            @RequestParam("sheet2Name") String sheet2Name,
            @RequestParam("startDate") String startDate,
            @RequestParam("endDate") String endDate) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            CwgjxBijiaoService.CompareResult compareResult = cwgjxBijiaoService.compare(
                    file1, sheet1Name, file2, sheet2Name, startDate, endDate, uploadDir);
            result.put("success", true);
            result.put("message", "数据比较完成");
            result.put("fileName", compareResult.fileName);
            CwgjxBijiaoService.ComparePreview preview = compareResult.preview;
            result.put("preview", preview);
            result.put("previewRows", compareResult.previewRows);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }
}
