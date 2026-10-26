package com.example.project.controller.finance.jxkh;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.jxkh.JxkhHbbgCustomMergeEngine;
import com.example.project.service.finance.jxkh.JxkhHbbgService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
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
     * 异步化（多文件）：请求线程内固化文件列表，提交 taskId，后台线程池执行
     */
    @PostMapping("/jxkh/hbbg")
    public Map<String, Object> jxkhHbbg(
            @RequestParam("files") List<MultipartFile> files,
            @RequestParam("targetCol") String targetCol,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            List<MultipartFile> safe = pinAll(files);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "jxkh/hbbg", "合并表格",
                    files == null || files.isEmpty() ? null : files.get(0).getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        String outputPath = jxkhHbbgService.processAndMerge(safe, targetCol, uploadDir);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "合并表格完成");
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

    /** 把 MultipartFile 列表固化为内存副本列表（异步线程安全） */
    protected List<MultipartFile> pinAll(List<MultipartFile> files) throws IOException {
        List<MultipartFile> out = new ArrayList<>();
        if (files != null) {
            for (MultipartFile f : files) {
                if (f == null || f.isEmpty()) continue;
                out.add(pin(f));
            }
        }
        return out;
    }

    /**
     * POST /api/finance/jxkh/hbbg-generic
     * 功能：通用合并表格 — 原样堆叠，不填月份（对齐 hbbg_zhf.py 通用合并模式）
     * 异步化（多文件）：提交即返回 taskId，后台线程池执行
     */
    @PostMapping("/jxkh/hbbg-generic")
    public Map<String, Object> jxkhHbbgGeneric(
            @RequestParam("files") List<MultipartFile> files,
            @RequestParam(value = "gapRows", defaultValue = "0") int gapRows,
            @RequestParam(value = "keepHeaderFile", defaultValue = "") String keepHeaderFile,
            @RequestParam(value = "skipRows", defaultValue = "1") int skipRows,
            @RequestParam(value = "shiftRules", defaultValue = "") String shiftRulesJson,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            // shiftRules 解析失败则同步返回错误
            List<JxkhHbbgService.ShiftRule> shiftRules = new ArrayList<>();
            if (shiftRulesJson != null && !shiftRulesJson.isEmpty()) {
                try {
                    shiftRules = parseShiftRules(shiftRulesJson);
                } catch (Exception e) {
                    Map<String, Object> err = new HashMap<>();
                    err.put("success", false);
                    err.put("message", "列移动规则解析失败：" + e.getMessage());
                    return err;
                }
            }
            final List<JxkhHbbgService.ShiftRule> rules = shiftRules;
            List<MultipartFile> safe = pinAll(files);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "jxkh/hbbg-generic", "通用合并表格",
                    files == null || files.isEmpty() ? null : files.get(0).getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        String outputPath = jxkhHbbgService.genericMerge(
                                safe, uploadDir, gapRows, keepHeaderFile, skipRows, rules);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "通用合并完成");
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

    /**
     * POST /api/finance/jxkh/hbbg-sheets
     * 功能：列出上传文件中所有工作表（用于多表合并模式的工作表勾选）
     */
    @PostMapping("/jxkh/hbbg-sheets")
    public Map<String, Object> jxkhHbbgSheets(@RequestParam("files") List<MultipartFile> files) {
        Map<String, Object> result = new HashMap<>();
        try {
            Map<String, List<JxkhHbbgCustomMergeEngine.SheetInfo>> sheetMap = new LinkedHashMap<>();
            for (MultipartFile file : files) {
                List<JxkhHbbgCustomMergeEngine.SheetInfo> sheets = JxkhHbbgCustomMergeEngine.listSheets(file);
                String name = file.getOriginalFilename() != null ? file.getOriginalFilename() : "unknown";
                sheetMap.put(name, sheets);
            }
            result.put("success", true);
            result.put("sheets", sheetMap);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "读取工作表列表失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/jxkh/hbbg-custom
     * 功能：多表合并（删列+公司命名），纯数据输出
     * 异步化（多文件）：提交即返回 taskId，后台线程池执行
     */
    @PostMapping("/jxkh/hbbg-custom")
    public Map<String, Object> jxkhHbbgCustom(
            @RequestParam("files") List<MultipartFile> files,
            @RequestParam("sheetSelections") String sheetSelectionsJson,
            @RequestParam("deleteColLetter") String deleteColLetter,
            @RequestParam("sheetNameColLetter") String sheetNameColLetter,
            @RequestParam(value = "headerRows", defaultValue = "5") int headerRows,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            com.fasterxml.jackson.databind.ObjectMapper mapper = new com.fasterxml.jackson.databind.ObjectMapper();
            List<JxkhHbbgCustomMergeEngine.SheetSelection> selections = mapper.readValue(
                sheetSelectionsJson,
                new com.fasterxml.jackson.core.type.TypeReference<List<JxkhHbbgCustomMergeEngine.SheetSelection>>() {}
            );
            List<MultipartFile> safe = pinAll(files);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "jxkh/hbbg-custom", "多表合并",
                    files == null || files.isEmpty() ? null : files.get(0).getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        String outputPath = jxkhHbbgService.customMerge(
                            safe, selections, deleteColLetter, sheetNameColLetter, headerRows, uploadDir
                        );
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "多表合并完成");
                        result.put("fileName", new File(outputPath).getName());
                        return result;
                    });
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件读取失败：" + e.getMessage());
            return err;
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "合并失败：" + e.getMessage());
            return err;
        }
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
