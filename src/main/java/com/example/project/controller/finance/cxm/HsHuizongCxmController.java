package com.example.project.controller.finance.cxm;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cxm.HsHuizongCxmService;
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
public class HsHuizongCxmController extends FinanceBaseController {

    @Autowired
    private HsHuizongCxmService hsHuizongCxmService;

    /**
     * POST /api/finance/summary
     * 功能：核算汇总（hs_huizong_cxm.py）
     * 根据备注2列缩写公司名，按供应商分组汇总打款金额
     */
    @PostMapping("/summary")
    public Map<String, Object> summary(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "types", required = false) String types,
            @RequestParam(value = "months", required = false) String months,
            @RequestParam(value = "amountCol", defaultValue = "-1") int amountCol,
            @RequestParam(value = "sheetName", required = false) String sheetName) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            String outputPath = hsHuizongCxmService.processSummary(file, types, months, amountCol, sheetName, uploadDir);
            result.put("success", true);
            result.put("message", "核算汇总完成");
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
     */
    @PostMapping("/summary/append")
    public Map<String, Object> summaryAppend(
            @RequestParam(value = "summaryFile", required = false) MultipartFile summaryFile,
            @RequestParam(value = "summaryFileName", required = false) String summaryFileName,
            @RequestParam(value = "checkFile", required = false) MultipartFile checkFile,
            @RequestParam(value = "checkFileName", required = false) String checkFileName,
            @RequestParam("targetFile") MultipartFile targetFile) {
        Map<String, Object> result = new HashMap<>();
        File summaryTemp = null;
        File checkTemp = null;
        File targetTemp = null;
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 处理汇总文件（可能直接传文件，也可能传文件名）
            String summaryPath = null;
            if (summaryFile != null && !summaryFile.isEmpty()) {
                summaryTemp = new File(dir, UUID.randomUUID() + "_summary.xlsx");
                try (InputStream in = summaryFile.getInputStream();
                     OutputStream out = new FileOutputStream(summaryTemp)) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
                }
                summaryPath = summaryTemp.getAbsolutePath();
            } else if (summaryFileName != null && !summaryFileName.isEmpty()) {
                File f = new File(dir, summaryFileName);
                if (f.exists()) summaryPath = f.getAbsolutePath();
            }

            // 处理核对结果文件
            String checkPath = null;
            if (checkFile != null && !checkFile.isEmpty()) {
                checkTemp = new File(dir, UUID.randomUUID() + "_check.xlsx");
                try (InputStream in = checkFile.getInputStream();
                     OutputStream out = new FileOutputStream(checkTemp)) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
                }
                checkPath = checkTemp.getAbsolutePath();
            } else if (checkFileName != null && !checkFileName.isEmpty()) {
                File f = new File(dir, checkFileName);
                if (f.exists()) checkPath = f.getAbsolutePath();
            }

            // 处理目标文件
            targetTemp = new File(dir, UUID.randomUUID() + "_target.xlsx");
            try (InputStream in = targetFile.getInputStream();
                 OutputStream out = new FileOutputStream(targetTemp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            String targetPath = targetTemp.getAbsolutePath();

            // 执行追加
            Map<String, Object> appendResult = hsHuizongCxmService.appendToTarget(summaryPath, checkPath, targetPath);
            String resultPath = (String) appendResult.get("targetPath");
            @SuppressWarnings("unchecked")
            List<String> appended = (List<String>) appendResult.get("appended");
            @SuppressWarnings("unchecked")
            List<String> skipped = (List<String>) appendResult.get("skipped");

            // 将结果重命名为用户上传的原始文件名（保留在 uploadDir 供下载）
            String originalName = targetFile.getOriginalFilename();
            if (originalName == null || originalName.isEmpty()) {
                originalName = "追加结果.xlsx";
            }
            File outputFile = new File(dir, originalName);
            if (outputFile.exists()) outputFile.delete();
            new File(resultPath).renameTo(outputFile);

            result.put("success", true);
            result.put("message", appendResult.get("message"));
            result.put("fileName", outputFile.getName());
            result.put("appended", appended);   // 前端可用来判断是否需要更新UI
            result.put("skipped", skipped);      // 前端可用来显示跳过提示
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "追加失败：" + e.getMessage());
        } finally {
            if (summaryTemp != null && summaryTemp.exists()) summaryTemp.delete();
            if (checkTemp != null && checkTemp.exists()) checkTemp.delete();
            // targetTemp 已被重命名为 outputFile，无需删除
        }
        return result;
    }

    /**
     * POST /api/finance/summary/batch
     * 功能：批量处理多个工作表，并合并结果
     * 对应 Python: hs_gui_actions.py::_process_data() 中多Sheet处理部分
     */
    @PostMapping("/summary/batch")
    public Map<String, Object> summaryBatch(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetNames") String sheetNames,
            @RequestParam(value = "types", required = false) String types,
            @RequestParam(value = "months", required = false) String months,
            @RequestParam(value = "amountCol", defaultValue = "-1") int amountCol) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 解析工作表名称列表
            List<String> sheetNameList = Arrays.asList(sheetNames.split(","));

            // 执行批量处理
            String outputPath = hsHuizongCxmService.processSummaryBatch(
                    file, sheetNameList, types, months, amountCol, uploadDir);

            result.put("success", true);
            result.put("message", "批量处理完成，共处理 " + sheetNameList.size() + " 个工作表");
            result.put("fileName", new File(outputPath).getName());

            // 返回预览数据（前50行）
            Map<String, Object> preview = readExcelPreview(new File(outputPath), 50);
            result.put("previewColumns", preview.get("columns"));
            result.put("previewData", preview.get("data"));
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "批量处理失败：" + e.getMessage());
        }
        return result;
    }
}
