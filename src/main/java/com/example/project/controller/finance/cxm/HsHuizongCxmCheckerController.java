package com.example.project.controller.finance.cxm;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cxm.CheckProcessor;
import com.example.project.service.finance.cxm.HsHuizongCxmChecker;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

@RestController
@RequestMapping("/api/finance")
public class HsHuizongCxmCheckerController extends FinanceBaseController {

    @Autowired
    private CheckProcessor checkProcessor;

    @Autowired
    private HsHuizongCxmChecker hsHuizongCxmChecker;

    /**
     * POST /api/finance/summary/check/prepare
     * 前端第一步：上传待核对文件，获取月份和介绍人列表供用户选择
     */
    @PostMapping("/summary/check/prepare")
    public Map<String, Object> summaryCheckPrepare(
            @RequestParam("checkFile") MultipartFile checkFile,
            @RequestParam("checkSheet") String checkSheet) {
        try {
            return hsHuizongCxmChecker.detectCheckColumns(checkFile, checkSheet);
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "检测失败：" + e.getMessage());
            return err;
        }
    }

    /**
     * POST /api/finance/check/detect-params
     * 功能：检测发放记录表参数（列名、月份、类型）
     * 前端选择文件+工作表后调用，返回下拉选项
     * 支持传入多个sheetName以合并所有表的月份
     */
    @PostMapping("/check/detect-params")
    public Map<String, Object> checkDetectParams(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetNames") List<String> sheetNames) {
        Map<String, Object> result = new HashMap<>();
        try {
            if (sheetNames == null || sheetNames.isEmpty()) {
                result.put("success", false);
                result.put("message", "请选择至少一个工作表");
                return result;
            }
            Map<String, Object> params = checkProcessor.detectDisbursementParams(file, sheetNames);
            result.putAll(params);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "参数检测失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/check/disbursement
     * 功能：汇总发放记录表（核对Tab - 区域A）
     * 供应商名称列自动检测，无需传入colC
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/check/disbursement")
    public Map<String, Object> checkDisbursement(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("colH") int colH,
            @RequestParam(value = "selectedMonths", required = false) String selectedMonths,
            @RequestParam(value = "selectedTypes", required = false) String selectedTypes,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "check/disbursement", "发放记录汇总",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = checkProcessor.processDisbursement(safeFile, sheetName, colH,
                            selectedMonths, selectedTypes, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "发放记录表汇总完成");
                    result.put("fileName", new File(outputPath).getName());
                    Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
                    result.put("previewColumns", preview.get("columns"));
                    result.put("previewData", preview.get("data"));
                    return result;
                });
    }

    /**
     * POST /api/finance/check/statement
     * 功能：汇总对账单（核对Tab - 区域B）
     */
    @PostMapping("/check/statement")
    public Map<String, Object> checkStatement(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam(value = "selectedMonths", required = false) String selectedMonths,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "check/statement", "对账单汇总",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = checkProcessor.processStatement(safeFile, sheetName, selectedMonths, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "对账单汇总完成");
                    result.put("fileName", new File(outputPath).getName());
                    Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
                    result.put("previewColumns", preview.get("columns"));
                    result.put("previewData", preview.get("data"));
                    return result;
                });
    }

    /**
     * POST /api/finance/check/reconcile
     * 功能：执行核对对比（核对Tab - 开始核对按钮）
     * 支持两种方式：
     * 1. 上传文件（disbursementFile, statementFile）
     * 2. 通过文件名从服务器获取（disbursementFileName, statementFileName）
     * 异步化：请求线程内解析两输入路径，提交 taskId，后台线程池执行核对
     */
    @PostMapping("/check/reconcile")
    public Map<String, Object> checkReconcile(
            @RequestParam(value = "disbursementFile", required = false) MultipartFile disbursementFile,
            @RequestParam(value = "statementFile", required = false) MultipartFile statementFile,
            @RequestParam(value = "disbursementFileName", required = false) String disbursementFileName,
            @RequestParam(value = "statementFileName", required = false) String statementFileName,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 请求线程内解析/落盘两输入文件（异步线程不能碰 MultipartFile）
            String disbursementPath = resolveInputPath(disbursementFile, disbursementFileName, dir, "_disbursement");
            if (disbursementPath == null) {
                Map<String, Object> err = new HashMap<>();
                err.put("success", false);
                err.put("message", "请上传发放记录汇总表或提供文件名");
                return err;
            }
            String statementPath = resolveInputPath(statementFile, statementFileName, dir, "_statement");
            if (statementPath == null) {
                Map<String, Object> err = new HashMap<>();
                err.put("success", false);
                err.put("message", "请上传对账单汇总表或提供文件名");
                return err;
            }
            final String disp = disbursementPath;
            final String stmt = statementPath;
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "check/reconcile", "执行核对",
                    null, () -> {
                        String outputPath = checkProcessor.reconcile(disp, stmt, uploadDir);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "核对完成");
                        result.put("fileName", new File(outputPath).getName());
                        Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
                        result.put("previewColumns", preview.get("columns"));
                        result.put("previewData", preview.get("data"));
                        return result;
                    });
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "核对失败：" + e.getMessage());
            return err;
        }
    }

    /** 上传文件→落盘临时文件取绝对路径；否则 fileName→uploads 下已有文件；都无返回 null */
    private String resolveInputPath(MultipartFile file, String fileName, File dir, String tag) throws Exception {
        if (file != null && !file.isEmpty()) {
            File temp = new File(dir, UUID.randomUUID() + tag + ".xlsx");
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(temp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            return temp.getAbsolutePath();
        }
        if (fileName != null && !fileName.isEmpty()) {
            File f = new File(uploadDir, fileName);
            if (f.exists()) return f.getAbsolutePath();
            throw new IllegalArgumentException("文件不存在：" + fileName);
        }
        return null;
    }

    /**
     * POST /api/finance/summary/check
     * 前端第二步：传入汇总结果文件名和待核对文件，传入用户选择的月份和介绍人，执行核对
     * 异步化：提交即返回 taskId，后台线程池执行
     */
    @PostMapping("/summary/check")
    public Map<String, Object> summaryCheck(
            @RequestParam("refFileName") String refFileName,
            @RequestParam("checkFile") MultipartFile checkFile,
            @RequestParam("checkSheet") String checkSheet,
            @RequestParam("selectedMonth") String selectedMonth,
            @RequestParam("selectedPersons") String selectedPersons,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            // 请求线程内固化 checkFile（异步线程安全）
            com.example.project.task.MemoryMultipartFile check = pin(checkFile);
            List<String> persons = Arrays.asList(selectedPersons.split(","));
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "summary/check", "汇总核对",
                    checkFile.getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        File refFile = new File(uploadDir, refFileName);
                        if (!refFile.exists()) {
                            throw new IllegalArgumentException("基准表文件不存在：" + refFileName);
                        }
                        hsHuizongCxmChecker.loadReference(refFile);
                        String outputPath = hsHuizongCxmChecker.executeCheck(
                                check, checkSheet, selectedMonth, persons, dir.getAbsolutePath());
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "核对完成");
                        result.put("fileName", new File(outputPath).getName());
                        Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
                        result.put("previewColumns", preview.get("columns"));
                        result.put("previewData", preview.get("data"));
                        return result;
                    });
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "核对失败：" + e.getMessage());
            return err;
        }
    }
}
