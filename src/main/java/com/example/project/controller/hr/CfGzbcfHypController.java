package com.example.project.controller.hr;

import com.example.project.service.hr.CfGzbcfHypService;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/hr")
public class CfGzbcfHypController extends HrBaseController {

    @Autowired
    private CfGzbcfHypService cfGzbcfHypService;

    /**
     * POST /api/hr/split-salary/sheets
     * 获取 Excel 文件的工作表列表
     */
    @PostMapping("/split-salary/sheets")
    public Map<String, Object> getSplitSalarySheets(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> sheets = cfGzbcfHypService.getSheetNames(file);
            result.put("success", true);
            result.put("sheets", sheets);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/split-salary/headers
     * 获取指定工作表第2行（表头行）的字段列表
     * headerRow 为 1-based，默认传 2 即可
     */
    @PostMapping("/split-salary/headers")
    public Map<String, Object> getSplitSalaryHeaders(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam(value = "headerRow", defaultValue = "2") int headerRow) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> headers = cfGzbcfHypService.getHeaders(file, sheetName, headerRow);
            result.put("success", true);
            result.put("headers", headers);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/split-salary/values
     * 获取指定列的唯一切片值
     */
    @PostMapping("/split-salary/values")
    public Map<String, Object> getSplitSalaryValues(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("colName") String colName) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> values = cfGzbcfHypService.getColumnValues(file, sheetName, colName);
            result.put("success", true);
            result.put("values", values);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/split-salary
     * 功能：工资表拆分（cf_gzbcf_hyp.py）
     * 按指定列拆分工资表，保留格式和公式
     * 支持自定义分组（groups 参数），不传则按值各自拆分
     */
    @SuppressWarnings("unchecked")
    @PostMapping("/split-salary")
    public Map<String, Object> splitSalary(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("splitColName") String splitColName,
            @RequestParam(value = "groups", required = false, defaultValue = "") String groupsJson) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String originalName = file.getOriginalFilename();
            String outputPath;
            if (groupsJson != null && !groupsJson.trim().isEmpty()) {
                // 解析 groups JSON，格式：[[\"IT部\",\"市场部\"],[\"行政部\"]]
                List<List<String>> groups = parseGroupsJson(groupsJson);
                outputPath = cfGzbcfHypService.splitSalaryWithGroups(file, originalName, sheetName, splitColName, groups, uploadDir);
            } else {
                outputPath = cfGzbcfHypService.splitSalary(file, originalName, sheetName, splitColName, uploadDir);
            }
            result.put("success", true);
            result.put("message", "工资表拆分完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    private final ObjectMapper objectMapper = new ObjectMapper();

    /**
     * 解析 groups JSON（Base64 + Jackson），格式：[[\"a\",\"b\"],[\"c\"]]
     */
    private List<List<String>> parseGroupsJson(String encoded) {
        List<List<String>> groups = new ArrayList<>();
        if (encoded == null || encoded.trim().isEmpty()) return groups;
        try {
            // Base64 解码
            byte[] decoded = Base64.getDecoder().decode(encoded);
            String json = new String(decoded, StandardCharsets.UTF_8);
            groups = objectMapper.readValue(json, new TypeReference<List<List<String>>>() {});
        } catch (Exception e) {
            // 解析失败返回空，让 Service 走普通拆分
        }
        return groups;
    }
}
