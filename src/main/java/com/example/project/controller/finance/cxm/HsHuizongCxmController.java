package com.example.project.controller.finance.cxm;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cxm.HsHuizongCxmService;
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
public class HsHuizongCxmController extends FinanceBaseController {

    @Autowired
    private HsHuizongCxmService hsHuizongCxmService;

    /**
     * POST /api/finance/summary
     * 功能：核算汇总（hs_huizong_cxm.py）
     * 根据备注2列缩写公司名，按供应商分组汇总打款金额
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/summary")
    public Map<String, Object> summary(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "types", required = false) String types,
            @RequestParam(value = "months", required = false) String months,
            @RequestParam(value = "amountCol", defaultValue = "-1") int amountCol,
            @RequestParam(value = "sheetName", required = false) String sheetName,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "summary", "核算汇总",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    String outputPath = hsHuizongCxmService.processSummary(safeFile, types, months, amountCol, sheetName, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "核算汇总完成");
                    result.put("fileName", new File(outputPath).getName());
                    Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
                    result.put("previewColumns", preview.get("columns"));
                    result.put("previewData", preview.get("data"));
                    return result;
                });
    }

    /**
     * POST /api/finance/summary/detect
     * 功能：检测 Excel 中的列信息（打款类型、月份、打款金额列）
     */
    @PostMapping("/summary/detect")
    public Map<String, Object> detectColumns(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "sheetName", required = false) String sheetName) {
        try {
            return hsHuizongCxmService.detectColumns(file, sheetName);
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "检测失败：" + e.getMessage());
            return err;
        }
    }

    /**
     * POST /api/finance/summary/batch-detect
     * 功能：批量检测多个工作表的列信息（预扫描）
     * 对应 Python: hs_gui_actions.py::_process_data() 中的预扫描逻辑
     */
    @PostMapping("/summary/batch-detect")
    public Map<String, Object> summaryBatchDetect(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetNames") String sheetNames) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> sheetNameList = Arrays.asList(sheetNames.split(","));
            Map<String, Object> detectResult = hsHuizongCxmService.detectColumnsBatch(file, sheetNameList);
            result.putAll(detectResult);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "批量检测失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/summary/append
     * 功能：将汇总结果和核对结果追加到目标文件
     * 目标文件不存在则创建，已存在则在原文件中追加新工作表
     * 异步化：请求线程内落盘输入文件，提交 taskId，后台线程池执行追加
     */
    @PostMapping("/summary/append")
    public Map<String, Object> summaryAppend(
            @RequestParam(value = "summaryFile", required = false) MultipartFile summaryFile,
            @RequestParam(value = "summaryFileName", required = false) String summaryFileName,
            @RequestParam(value = "checkFile", required = false) MultipartFile checkFile,
            @RequestParam(value = "checkFileName", required = false) String checkFileName,
            @RequestParam("targetFile") MultipartFile targetFile,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            final String originalName = (targetFile.getOriginalFilename() == null || targetFile.getOriginalFilename().isEmpty())
                    ? "追加结果.xlsx" : targetFile.getOriginalFilename();

            // 请求线程内落盘输入文件（异步线程不能碰 MultipartFile）
            final String summaryPath = persistInput(summaryFile, summaryFileName, dir, "summary");
            final String checkPath = persistInput(checkFile, checkFileName, dir, "check");
            final File targetTemp = new File(dir, UUID.randomUUID() + "_target.xlsx");
            try (InputStream in = targetFile.getInputStream();
                 OutputStream out = new FileOutputStream(targetTemp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "summary/append", "汇总结果追加",
                    originalName,
                    () -> {
                        Map<String, Object> appendResult = hsHuizongCxmService.appendToTarget(
                                summaryPath, checkPath, targetTemp.getAbsolutePath());
                        String resultPath = (String) appendResult.get("targetPath");
                        @SuppressWarnings("unchecked")
                        List<String> appended = (List<String>) appendResult.get("appended");
                        @SuppressWarnings("unchecked")
                        List<String> skipped = (List<String>) appendResult.get("skipped");
                        File outputFile = new File(dir, originalName);
                        if (outputFile.exists()) outputFile.delete();
                        new File(resultPath).renameTo(outputFile);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", appendResult.get("message"));
                        result.put("fileName", outputFile.getName());
                        result.put("appended", appended);
                        result.put("skipped", skipped);
                        return result;
                    });
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "追加失败：" + e.getMessage());
            return err;
        }
    }

    /** 把上传文件或已有文件名落盘/解析为绝对路径；二者皆无返回 null */
    private String persistInput(MultipartFile file, String fileName, File dir, String tag) throws Exception {
        if (file != null && !file.isEmpty()) {
            File temp = new File(dir, UUID.randomUUID() + "_" + tag + ".xlsx");
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(temp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            return temp.getAbsolutePath();
        }
        if (fileName != null && !fileName.isEmpty()) {
            File f = new File(dir, fileName);
            if (f.exists()) return f.getAbsolutePath();
        }
        return null;
    }

    /**
     * POST /api/finance/summary/batch
     * 功能：批量处理多个工作表，并合并结果
     * 对应 Python: hs_gui_actions.py::_process_data() 中多Sheet处理部分
     * 异步化：提交即返回 taskId，后台线程池执行，前端轮询 /api/tasks/{taskId}
     */
    @PostMapping("/summary/batch")
    public Map<String, Object> summaryBatch(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetNames") String sheetNames,
            @RequestParam(value = "types", required = false) String types,
            @RequestParam(value = "months", required = false) String months,
            @RequestParam(value = "amountCol", defaultValue = "-1") int amountCol,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        return doSubmitAsync(file, deviceId(deviceIdHeader), "finance", "summary/batch", "核算汇总(批量)",
                safeFile -> {
                    File dir = new File(uploadDir);
                    if (!dir.exists()) dir.mkdirs();
                    List<String> sheetNameList = Arrays.asList(sheetNames.split(","));
                    String outputPath = hsHuizongCxmService.processSummaryBatch(
                            safeFile, sheetNameList, types, months, amountCol, uploadDir);
                    Map<String, Object> result = new HashMap<>();
                    result.put("message", "批量处理完成，共处理 " + sheetNameList.size() + " 个工作表");
                    result.put("fileName", new File(outputPath).getName());
                    Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
                    result.put("previewColumns", preview.get("columns"));
                    result.put("previewData", preview.get("data"));
                    return result;
                });
    }
}
