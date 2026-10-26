package com.example.project.controller.hr;

import com.example.project.service.hr.HrGzbPipeiService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/hr")
public class HrGzbPipeiController extends HrBaseController {

    @Autowired
    private HrGzbPipeiService hrGzbPipeiService;

    /**
     * POST /api/hr/gzb-pipei/sheets
     * 获取 Excel 文件的工作表列表（秒级辅助接口，保持同步）
     */
    @PostMapping("/gzb-pipei/sheets")
    public Map<String, Object> getSheetNames(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> sheets = hrGzbPipeiService.getSheetNames(file);
            result.put("success", true);
            result.put("sheets", sheets);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/gzb-pipei/preview
     * 预览指定工作表的前3行（秒级辅助接口，保持同步）
     */
    @PostMapping("/gzb-pipei/preview")
    public Map<String, Object> previewRows(@RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<HrGzbPipeiService.RowPreview> rows = hrGzbPipeiService.previewRows(file, sheetName, 3, 30);
            result.put("success", true);
            result.put("rows", rows);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/gzb-pipei/headers
     * 获取指定表头行的字段列表（秒级辅助接口，保持同步）
     */
    @PostMapping("/gzb-pipei/headers")
    public Map<String, Object> getHeaders(@RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("headerRow") int headerRow) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> headers = hrGzbPipeiService.getHeaders(file, sheetName, headerRow);
            result.put("success", true);
            result.put("headers", headers);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/gzb-pipei/mode-a
     * 模式A：甲方工资表匹配内部工资表（双文件，异步化）
     */
    @PostMapping("/gzb-pipei/mode-a")
    public Map<String, Object> gzbPipeiModeA(
            @RequestParam("jiaFile") MultipartFile jiaFile,
            @RequestParam("neiFile") MultipartFile neiFile,
            @RequestParam("jiaSheet") String jiaSheet,
            @RequestParam("neiSheet") String neiSheet,
            @RequestParam("jiaHeaderRow") int jiaHeaderRow,
            @RequestParam("neiHeaderRow") int neiHeaderRow,
            @RequestParam("jiaKey") String jiaKey,
            @RequestParam("neiKey") String neiKey,
            @RequestParam("fieldMapping") String fieldMapping,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            if (jiaKey == null || jiaKey.trim().isEmpty()) throw new IllegalArgumentException("请选择甲方关键列");
            if (neiKey == null || neiKey.trim().isEmpty()) throw new IllegalArgumentException("请选择内部关键列");
            com.example.project.task.MemoryMultipartFile jia = pin(jiaFile);
            com.example.project.task.MemoryMultipartFile nei = pin(neiFile);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "hr", "gzb-pipei/mode-a", "工资表匹配(模式A)",
                    jiaFile.getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        HrGzbPipeiService.ProcessResult pr = hrGzbPipeiService.processModeA(
                                jia, nei, jiaSheet, neiSheet,
                                jiaHeaderRow, neiHeaderRow, jiaKey.trim(), neiKey.trim(), fieldMapping, uploadDir);
                        Map<String, Object> result = new HashMap<>();
                        putResult(result, pr);
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
     * POST /api/hr/gzb-pipei/mode-b
     * 模式B：同一文件内两个工作表互相匹配（单文件，异步化）
     */
    @PostMapping("/gzb-pipei/mode-b")
    public Map<String, Object> gzbPipeiModeB(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetA") String sheetA,
            @RequestParam("sheetB") String sheetB,
            @RequestParam("headerA") int headerA,
            @RequestParam("headerB") int headerB,
            @RequestParam("keyA") String keyA,
            @RequestParam("keyB") String keyB,
            @RequestParam("fieldMapping") String fieldMapping,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "hr", "gzb-pipei/mode-b", "工资表匹配(模式B)",
                safeFile -> {
                    if (keyA == null || keyA.trim().isEmpty()) throw new IllegalArgumentException("请选择工作簿A关键列");
                    if (keyB == null || keyB.trim().isEmpty()) throw new IllegalArgumentException("请选择工作簿B关键列");
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    HrGzbPipeiService.ProcessResult pr = hrGzbPipeiService.processModeB(
                            safeFile, sheetA, sheetB, headerA, headerB, keyA.trim(), keyB.trim(), fieldMapping, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    putResult(result, pr);
                    return result;
                });
    }

    private void putResult(Map<String, Object> result, HrGzbPipeiService.ProcessResult pr) {
        result.put("success", true);
        result.put("message", "工资表匹配完成");
        result.put("fileName", new File(pr.outputPath).getName());
        result.put("compareRows", pr.matchedCount);
        result.put("diffCount", pr.diffCount);
        result.put("unmatchedKeys", pr.unmatchedKeys);
        result.put("totalInfo", pr.totalInfo);
    }
}
