package com.example.project.controller.finance.cxm;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cxm.CheckProcessor;
import com.example.project.service.finance.cxm.HsHuizongCxmChecker;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
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
     */
    @PostMapping("/check/disbursement")
    public Map<String, Object> checkDisbursement(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("colH") int colH,
            @RequestParam(value = "selectedMonths", required = false) String selectedMonths,
            @RequestParam(value = "selectedTypes", required = false) String selectedTypes) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = checkProcessor.processDisbursement(file, sheetName, colH,
                    selectedMonths, selectedTypes, uploadDir);
            result.put("success", true);
            result.put("message", "发放记录表汇总完成");
            result.put("fileName", new File(outputPath).getName());

            // 返回预览数据（前50行）
            Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
            result.put("previewColumns", preview.get("columns"));
            result.put("previewData", preview.get("data"));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/check/statement
     * 功能：汇总对账单（核对Tab - 区域B）
     */
    @PostMapping("/check/statement")
    public Map<String, Object> checkStatement(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam(value = "selectedMonths", required = false) String selectedMonths) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = checkProcessor.processStatement(file, sheetName, selectedMonths, uploadDir);
            result.put("success", true);
            result.put("message", "对账单汇总完成");
            result.put("fileName", new File(outputPath).getName());

            // 返回预览数据（前50行）
            Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
            result.put("previewColumns", preview.get("columns"));
            result.put("previewData", preview.get("data"));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/check/reconcile
     * 功能：执行核对对比（核对Tab - 开始核对按钮）
     * 支持两种方式：
     * 1. 上传文件（disbursementFile, statementFile）
     * 2. 通过文件名从服务器获取（disbursementFileName, statementFileName）
     */
    @PostMapping("/check/reconcile")
    public Map<String, Object> checkReconcile(
            @RequestParam(value = "disbursementFile", required = false) MultipartFile disbursementFile,
            @RequestParam(value = "statementFile", required = false) MultipartFile statementFile,
            @RequestParam(value = "disbursementFileName", required = false) String disbursementFileName,
            @RequestParam(value = "statementFileName", required = false) String statementFileName) {
        Map<String, Object> result = new HashMap<>();
        File disbursementTemp = null;
        File statementTemp = null;
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 处理发放记录汇总表
            String disbursementPath = null;
            if (disbursementFile != null && !disbursementFile.isEmpty()) {
                // 方式1：上传文件
                disbursementTemp = new File(dir, UUID.randomUUID() + "_disbursement.xlsx");
                try (InputStream in = disbursementFile.getInputStream();
                     OutputStream out = new FileOutputStream(disbursementTemp)) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
                }
                disbursementPath = disbursementTemp.getAbsolutePath();
            } else if (disbursementFileName != null && !disbursementFileName.isEmpty()) {
                // 方式2：通过文件名从服务器获取
                File f = new File(uploadDir, disbursementFileName);
                if (f.exists()) {
                    disbursementPath = f.getAbsolutePath();
                } else {
                    result.put("success", false);
                    result.put("message", "发放记录汇总表文件不存在：" + disbursementFileName);
                    return result;
                }
            } else {
                result.put("success", false);
                result.put("message", "请上传发放记录汇总表或提供文件名");
                return result;
            }

            // 处理对账单汇总表
            String statementPath = null;
            if (statementFile != null && !statementFile.isEmpty()) {
                // 方式1：上传文件
                statementTemp = new File(dir, UUID.randomUUID() + "_statement.xlsx");
                try (InputStream in = statementFile.getInputStream();
                     OutputStream out = new FileOutputStream(statementTemp)) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
                }
                statementPath = statementTemp.getAbsolutePath();
            } else if (statementFileName != null && !statementFileName.isEmpty()) {
                // 方式2：通过文件名从服务器获取
                File f = new File(uploadDir, statementFileName);
                if (f.exists()) {
                    statementPath = f.getAbsolutePath();
                } else {
                    result.put("success", false);
                    result.put("message", "对账单汇总表文件不存在：" + statementFileName);
                    return result;
                }
            } else {
                result.put("success", false);
                result.put("message", "请上传对账单汇总表或提供文件名");
                return result;
            }

            // 执行核对
            String outputPath = checkProcessor.reconcile(
                    disbursementPath,
                    statementPath,
                    uploadDir);

            result.put("success", true);
            result.put("message", "核对完成");
            result.put("fileName", new File(outputPath).getName());

            // 返回预览数据（前50行）
            Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
            result.put("previewColumns", preview.get("columns"));
            result.put("previewData", preview.get("data"));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "核对失败：" + e.getMessage());
        } finally {
            // 清理临时文件（仅当使用上传方式时）
            if (disbursementTemp != null && disbursementTemp.exists()) disbursementTemp.delete();
            if (statementTemp != null && statementTemp.exists()) statementTemp.delete();
        }
        return result;
    }

    /**
     * POST /api/finance/summary/check
     * 前端第二步：传入汇总结果文件名和待核对文件，传入用户选择的月份和介绍人，执行核对
     */
    @PostMapping("/summary/check")
    public Map<String, Object> summaryCheck(
            @RequestParam("refFileName") String refFileName,
            @RequestParam("checkFile") MultipartFile checkFile,
            @RequestParam("checkSheet") String checkSheet,
            @RequestParam("selectedMonth") String selectedMonth,
            @RequestParam("selectedPersons") String selectedPersons) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 从 uploadDir 读取基准表文件
            File refFile = new File(uploadDir, refFileName);
            if (!refFile.exists()) {
                result.put("success", false);
                result.put("message", "基准表文件不存在：" + refFileName);
                return result;
            }

            // 加载基准表
            hsHuizongCxmChecker.loadReference(refFile);

            // 解析选择的介绍人
            List<String> persons = Arrays.asList(selectedPersons.split(","));

            // 执行核对
            String outputPath = hsHuizongCxmChecker.executeCheck(
                    checkFile, checkSheet, selectedMonth, persons,
                    dir.getAbsolutePath());
            result.put("success", true);
            result.put("message", "核对完成");
            result.put("fileName", new File(outputPath).getName());
            // 返回预览数据（前50行）
            Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
            result.put("previewColumns", preview.get("columns"));
            result.put("previewData", preview.get("data"));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "核对失败：" + e.getMessage());
        }
        return result;
    }
}
