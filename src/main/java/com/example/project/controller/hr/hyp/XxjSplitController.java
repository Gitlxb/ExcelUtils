package com.example.project.controller.hr.hyp;

import com.example.project.controller.hr.HrBaseController;
import com.example.project.service.hr.hyp.XxjSplitService;
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

/**
 * 孝心金拆分接口（对应桌面版 hr/hyp/cf_xxj_hyp_gui.py）。
 *
 * <p>与工资表拆分接口的差异：{@code /split-xxj/values} 返回的是「段」列表
 * （值 + 行范围 + 行数 + 金额合计），而不是去重后的值列表 —— 因为同名但不连续的两段
 * 需要分别展示与选择；组合拆分提交的 groups 是段序号（0-based）。
 */
@RestController
@RequestMapping("/api/hr")
public class XxjSplitController extends HrBaseController {

    @Autowired
    private XxjSplitService xxjSplitService;

    private final ObjectMapper objectMapper = new ObjectMapper();

    @PostMapping("/split-xxj/sheets")
    public Map<String, Object> getXxjSheets(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            result.put("success", true);
            result.put("sheets", xxjSplitService.getSheetNames(file));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    @PostMapping("/split-xxj/headers")
    public Map<String, Object> getXxjHeaders(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam(value = "headerRow", defaultValue = "2") int headerRow) {
        Map<String, Object> result = new HashMap<>();
        try {
            result.put("success", true);
            result.put("headers", xxjSplitService.getHeaders(file, sheetName, headerRow));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /** 识别拆分分组（段）：值 + 行范围 + 行数 + 金额合计 */
    @PostMapping("/split-xxj/values")
    public Map<String, Object> getXxjSegments(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("colName") String colName) {
        Map<String, Object> result = new HashMap<>();
        try {
            result.put("success", true);
            result.put("segments", xxjSplitService.getSegments(file, sheetName, colName));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    @PostMapping("/split-xxj")
    public Map<String, Object> splitXxj(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("splitColName") String splitColName,
            @RequestParam(value = "groups", required = false, defaultValue = "") String groupsJson,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "hr", "split-xxj", "孝心金拆分",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String originalName = safeFile.getOriginalFilename();
                    String outputPath;
                    List<List<Integer>> groups = parseGroupsJson(groupsJson);
                    if (!groups.isEmpty()) {
                        outputPath = xxjSplitService.splitXxjWithGroups(
                                safeFile, originalName, sheetName, splitColName, groups, uploadDir);
                    } else {
                        outputPath = xxjSplitService.splitXxj(
                                safeFile, originalName, sheetName, splitColName, uploadDir);
                    }
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "孝心金拆分完成");
                    result.put("fileName", new File(outputPath).getName());
                    return result;
                });
    }

    /** 解析 groups JSON（Base64 + Jackson），格式：[[0,1],[5]]（段序号） */
    private List<List<Integer>> parseGroupsJson(String encoded) {
        List<List<Integer>> groups = new ArrayList<>();
        if (encoded == null || encoded.trim().isEmpty()) return groups;
        try {
            byte[] decoded = Base64.getDecoder().decode(encoded);
            String json = new String(decoded, StandardCharsets.UTF_8);
            groups = objectMapper.readValue(json, new TypeReference<List<List<Integer>>>() { });
        } catch (Exception e) {
            // 解析失败返回空，让 Service 走普通拆分
        }
        return groups;
    }
}
