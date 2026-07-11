package com.example.project.controller.finance.jxkh;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.jxkh.JxkhHbbgService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/finance")
public class JxkhHbbgController extends FinanceBaseController {

    @Autowired
    private JxkhHbbgService jxkhHbbgService;

    /**
     * POST /api/finance/jxkh/hbbg
     * 功能：合并表格（frame_hbbg.py）
     */
    @PostMapping("/jxkh/hbbg")
    public Map<String, Object> jxkhHbbg(
            @RequestParam("files") List<MultipartFile> files,
            @RequestParam("targetCol") String targetCol) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = jxkhHbbgService.processAndMerge(files, targetCol, uploadDir);
            result.put("success", true);
            result.put("message", "合并表格完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/jxkh/hbbg-generic
     * 功能：通用合并表格 — 原样堆叠，不填月份（对齐 hbbg_zhf.py 通用合并模式）
     */
    @PostMapping("/jxkh/hbbg-generic")
    public Map<String, Object> jxkhHbbgGeneric(
            @RequestParam("files") List<MultipartFile> files,
            @RequestParam(value = "gapRows", defaultValue = "0") int gapRows,
            @RequestParam(value = "keepHeaderFile", defaultValue = "") String keepHeaderFile,
            @RequestParam(value = "skipRows", defaultValue = "1") int skipRows,
            @RequestParam(value = "shiftRules", defaultValue = "") String shiftRulesJson) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 解析列移动规则
            List<JxkhHbbgService.ShiftRule> shiftRules = new ArrayList<>();
            if (shiftRulesJson != null && !shiftRulesJson.isEmpty()) {
                try {
                    shiftRules = parseShiftRules(shiftRulesJson);
                } catch (Exception e) {
                    result.put("success", false);
                    result.put("message", "列移动规则解析失败：" + e.getMessage());
                    return result;
                }
            }

            String outputPath = jxkhHbbgService.genericMerge(files, uploadDir, gapRows, keepHeaderFile, skipRows, shiftRules);
            result.put("success", true);
            result.put("message", "通用合并完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "合并失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * 解析列移动规则 JSON
     */
    private List<JxkhHbbgService.ShiftRule> parseShiftRules(String json) throws Exception {
        List<JxkhHbbgService.ShiftRule> rules = new ArrayList<>();
        if (json == null || json.isEmpty() || "[]".equals(json)) {
            return rules;
        }

        com.fasterxml.jackson.databind.ObjectMapper mapper = new com.fasterxml.jackson.databind.ObjectMapper();
        List<Map<String, Object>> list = mapper.readValue(json, new com.fasterxml.jackson.core.type.TypeReference<List<Map<String, Object>>>() {});

        for (Map<String, Object> item : list) {
            // fileNames 是数组（支持多选）
            Object fileNamesObj = item.get("fileNames");
            List<String> fileNames = new ArrayList<>();
            if (fileNamesObj instanceof List) {
                for (Object f : (List<?>) fileNamesObj) {
                    fileNames.add(f.toString());
                }
            } else if (fileNamesObj instanceof String) {
                // 兼容旧格式（单文件字符串）
                fileNames.add((String) fileNamesObj);
            }
            if (fileNames.isEmpty()) continue;

            String colLetter = (String) item.get("colLetter");
            int colIndex = colLetterToIndex(colLetter);

            Object shiftCountObj = item.get("shiftCount");
            int shiftCount = 1;
            if (shiftCountObj instanceof Integer) {
                shiftCount = (Integer) shiftCountObj;
            } else if (shiftCountObj instanceof Double) {
                shiftCount = ((Double) shiftCountObj).intValue();
            }

            // 为每个文件创建一条规则
            for (String fileName : fileNames) {
                JxkhHbbgService.ShiftRule rule = new JxkhHbbgService.ShiftRule();
                rule.setFileName(fileName);
                rule.setColIndex(colIndex);
                rule.setShiftCount(shiftCount);
                rules.add(rule);
            }
        }
        return rules;
    }

    /**
     * 列字母转索引（A=0, B=1, ...）
     */
    private int colLetterToIndex(String colLetter) {
        if (colLetter == null || colLetter.isEmpty()) {
            return 0;
        }
        int result = 0;
        for (char c : colLetter.toUpperCase().toCharArray()) {
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1; // 0-based
    }
}
