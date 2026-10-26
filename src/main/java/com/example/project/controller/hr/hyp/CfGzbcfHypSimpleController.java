package com.example.project.controller.hr.hyp;

import com.example.project.service.hr.hyp.CfGzbcfHypSimpleService;
import com.example.project.controller.hr.HrBaseController;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
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
public class CfGzbcfHypSimpleController extends HrBaseController {

    @Autowired
    private CfGzbcfHypSimpleService cfGzbcfHypSimpleService;

    @PostMapping("/split-salary-simple/sheets")
    public Map<String, Object> getSplitSalarySimpleSheets(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            result.put("success", true);
            result.put("sheets", cfGzbcfHypSimpleService.getSheetNames(file));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    @PostMapping("/split-salary-simple/headers")
    public Map<String, Object> getSplitSalarySimpleHeaders(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam(value = "headerRow", defaultValue = "2") int headerRow) {
        Map<String, Object> result = new HashMap<>();
        try {
            result.put("success", true);
            result.put("headers", cfGzbcfHypSimpleService.getHeaders(file, sheetName, headerRow));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    @PostMapping("/split-salary-simple/values")
    public Map<String, Object> getSplitSalarySimpleValues(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("colName") String colName) {
        Map<String, Object> result = new HashMap<>();
        try {
            result.put("success", true);
            result.put("values", cfGzbcfHypSimpleService.getColumnValues(file, sheetName, colName));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    @SuppressWarnings("unchecked")
    @PostMapping("/split-salary-simple")
    public Map<String, Object> splitSalarySimple(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("splitColName") String splitColName,
            @RequestParam(value = "groups", required = false, defaultValue = "") String groupsJson,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "hr", "split-salary-simple", "工资表拆分(简化版)",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String originalName = safeFile.getOriginalFilename();
                    String outputPath;
                    if (groupsJson != null && !groupsJson.trim().isEmpty()) {
                        List<List<String>> groups = parseGroupsJson(groupsJson);
                        outputPath = cfGzbcfHypSimpleService.splitSalaryWithGroups(safeFile, originalName, sheetName, splitColName, groups, uploadDir);
                    } else {
                        outputPath = cfGzbcfHypSimpleService.splitSalary(safeFile, originalName, sheetName, splitColName, uploadDir);
                    }
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "工资表拆分完成（简化版）");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
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
