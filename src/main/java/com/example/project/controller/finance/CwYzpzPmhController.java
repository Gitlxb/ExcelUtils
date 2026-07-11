package com.example.project.controller.finance;

import com.example.project.service.finance.CwYzpzPmhService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/finance")
public class CwYzpzPmhController extends FinanceBaseController {

    @Autowired
    private CwYzpzPmhService cwYzpzPmhService;

    /**
     * POST /api/finance/yzpz/prepare
     * 功能：预支平账-准备阶段（cw_yzpz_pmh.py）
     */
    @PostMapping("/yzpz/prepare")
    public Map<String, Object> yzpzPrepare(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheet") String sheet,
            @RequestParam(value = "round", defaultValue = "1") int round) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwYzpzPmhService.prepare(file, sheet, round, uploadDir);
            result.put("success", true);
            result.put("message", "准备完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "准备失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/copy-names
     * 功能：预支平账-复制姓名
     */
    @PostMapping("/yzpz/copy-names")
    public Map<String, Object> yzpzCopyNames(
            @RequestParam("file") MultipartFile file,
            @RequestParam("checkFile") MultipartFile checkFile,
            @RequestParam("checkSheet") String checkSheet,
            @RequestParam(value = "round", defaultValue = "1") int round) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwYzpzPmhService.copyNames(file, checkFile, checkSheet, round, uploadDir);
            result.put("success", true);
            result.put("message", "姓名复制完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "复制失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/match
     * 功能：预支平账-姓名匹配
     */
    @PostMapping("/yzpz/match")
    public Map<String, Object> yzpzMatch(
            @RequestParam("file") MultipartFile file,
            @RequestParam("mergedSheet") String mergedSheet,
            @RequestParam(value = "units", required = false) String units,
            @RequestParam(value = "round", defaultValue = "1") int round) {
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            return cwYzpzPmhService.matchNames(file, mergedSheet, units, round, uploadDir);
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "匹配失败：" + e.getMessage());
            return err;
        }
    }

    /**
     * POST /api/finance/yzpz/process-deduction
     * 功能：预支平账-处理扣款数据
     */
    @PostMapping("/yzpz/process-deduction")
    public Map<String, Object> yzpzProcessDeduction(
            @RequestParam("file") MultipartFile file,
            @RequestParam("dedFile") MultipartFile dedFile,
            @RequestParam("dedSheet") String dedSheet,
            @RequestParam(value = "columnType", required = false) String columnType,
            @RequestParam(value = "round", defaultValue = "1") int round) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwYzpzPmhService.processDeduction(file, dedFile, dedSheet, columnType, round, uploadDir);
            result.put("success", true);
            result.put("message", "扣款数据处理完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/balance
     * 功能：预支平账-执行平账
     */
    @PostMapping("/yzpz/balance")
    public Map<String, Object> yzpzBalance(
            @RequestParam("file") MultipartFile file,
            @RequestParam("salarySheet") String salarySheet,
            @RequestParam("month") String month,
            @RequestParam(value = "round", defaultValue = "1") int round,
            @RequestParam(value = "units", required = false) String units) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwYzpzPmhService.balance(file, salarySheet, month, round, units, uploadDir);
            result.put("success", true);
            result.put("message", "平账完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "平账失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/detect-columns
     * 功能：检测扣款文件中的可用列类型（需扣回、预支扣款、预支扣回）
     */
    @PostMapping("/yzpz/detect-columns")
    public Map<String, Object> yzpzDetectColumns(
            @RequestParam("dedFile") MultipartFile dedFile,
            @RequestParam("dedSheet") String dedSheet) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> columns = cwYzpzPmhService.detectDeductionColumns(dedFile, dedSheet);
            result.put("success", true);
            result.put("columns", columns);
            result.put("message", "检测到 " + columns.size() + " 种列类型");
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "检测失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/detect-units
     * 功能：检测主文件工资表中单位列的唯一值
     */
    @PostMapping("/yzpz/detect-units")
    public Map<String, Object> yzpzDetectUnits(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheet") String sheet) {
        Map<String, Object> result = new HashMap<>();
        try {
            List<String> units = cwYzpzPmhService.detectUnits(file, sheet);
            result.put("success", true);
            result.put("units", units);
            result.put("message", "检测到 " + units.size() + " 个单位");
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "检测失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/prepare-and-check
     * 功能：合并准备+核对（prepare + copyNames + matchNames + processDeduction）
     * 前端步骤1"核对"调用此接口，两步合一
     */
    @PostMapping("/yzpz/prepare-and-check")
    public Map<String, Object> yzpzPrepareAndCheck(
            @RequestParam("file") MultipartFile file,
            @RequestParam("checkFile") MultipartFile checkFile,
            @RequestParam("dedFile") MultipartFile dedFile,
            @RequestParam("sheet") String sheet,
            @RequestParam("checkSheet") String checkSheet,
            @RequestParam("dedSheet") String dedSheet,
            @RequestParam(value = "units", required = false) String units,
            @RequestParam(value = "columnType", required = false) String columnType,
            @RequestParam(value = "round", defaultValue = "1") int round) {
        Map<String, Object> result = new HashMap<>();
        File tempFile = null;
        File afterPrepareFile = null;
        File afterCopyFile = null;
        File afterMatchFile = null;
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 第1步：prepare（创建核对工作表）
            String afterPrepare = cwYzpzPmhService.prepare(file, sheet, round, uploadDir);
            afterPrepareFile = new File(afterPrepare);
            tempFile = afterPrepareFile;

            // 第2步：copyNames（复制姓名）
            String afterCopy = cwYzpzPmhService.copyNames(
                    toMultipart(tempFile, file.getOriginalFilename()),
                    checkFile, checkSheet, round, uploadDir);
            afterCopyFile = new File(afterCopy);
            tempFile.delete();
            tempFile = afterCopyFile;

            // 第3步：matchNames（姓名匹配，多单位检测）
            Map<String, Object> matchResult = cwYzpzPmhService.matchNames(
                    toMultipart(tempFile, new File(afterCopy).getName()),
                    sheet, units, round, uploadDir);
            if (Boolean.FALSE.equals(matchResult.get("success"))) {
                result.put("success", false);
                result.put("message", matchResult.get("message"));
                return result;
            }
            // 多单位检测：需要前端选择单位
            if (Boolean.TRUE.equals(matchResult.get("needSelectUnits"))) {
                result.put("success", true);
                result.put("needSelectUnits", true);
                result.put("units", matchResult.get("units"));
                result.put("fileName", new File(afterCopy).getName());
                result.put("message", "请选择单位后继续");
                return result;
            }
            afterMatchFile = new File((String) matchResult.get("filePath"));
            tempFile.delete();
            tempFile = afterMatchFile;

            // 第4步：processDeduction（处理扣款数据）
            String afterDeduction = cwYzpzPmhService.processDeduction(
                    toMultipart(tempFile, new File((String) matchResult.get("filePath")).getName()),
                    dedFile, dedSheet, columnType, round, uploadDir);
            tempFile.delete();
            tempFile = null;

            File finalFile = new File(afterDeduction);
            result.put("success", true);
            result.put("message", "核对完成（准备+复制姓名+姓名匹配+处理扣款）");
            result.put("fileName", finalFile.getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "核对失败：" + e.getMessage());
            e.printStackTrace();
        } finally {
            // 清理临时文件
            if (afterPrepareFile != null && afterPrepareFile.exists()) afterPrepareFile.delete();
            if (afterCopyFile != null && afterCopyFile.exists()) afterCopyFile.delete();
            if (afterMatchFile != null && afterMatchFile.exists()) afterMatchFile.delete();
            if (tempFile != null && tempFile.exists()) tempFile.delete();
        }
        return result;
    }

    /**
     * POST /api/finance/yzpz/check
     * 功能：【核对】按钮链式调用（复制姓名 + 姓名匹配 + 处理扣款）
     * 将多个步骤合并为一个接口调用，前端只需点击一次【核对】按钮
     * @deprecated 已被 prepare-and-check 取代，保留以兼容旧前端
     */
    @PostMapping("/yzpz/check")
    public Map<String, Object> yzpzCheck(
            @RequestParam("file") MultipartFile file,
            @RequestParam("checkFile") MultipartFile checkFile,
            @RequestParam("dedFile") MultipartFile dedFile,
            @RequestParam("sheet") String sheet,
            @RequestParam("checkSheet") String checkSheet,
            @RequestParam("dedSheet") String dedSheet,
            @RequestParam(value = "units", required = false) String units,
            @RequestParam(value = "columnType", required = false) String columnType,
            @RequestParam(value = "round", defaultValue = "1") int round) {
        Map<String, Object> result = new HashMap<>();
        File tempFile = null;
        File afterCopyFile = null;
        File afterMatchFile = null;
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // 第1步：保存上传的文件到临时目录
            tempFile = File.createTempFile("yzpz_check_", ".xlsx", dir);
            file.transferTo(tempFile);

            // 第2步：复制姓名
            String afterCopy = cwYzpzPmhService.copyNames(
                    toMultipart(tempFile, file.getOriginalFilename()),
                    checkFile, checkSheet, round, uploadDir);
            afterCopyFile = new File(afterCopy);
            tempFile.delete();
            tempFile = afterCopyFile;

            // 第3步：姓名匹配（多单位检测）
            Map<String, Object> matchResult = cwYzpzPmhService.matchNames(
                    toMultipart(tempFile, new File(afterCopy).getName()),
                    sheet, units, round, uploadDir);
            if (Boolean.FALSE.equals(matchResult.get("success"))) {
                result.put("success", false);
                result.put("message", matchResult.get("message"));
                return result;
            }
            // 多单位检测：需要前端选择单位
            if (Boolean.TRUE.equals(matchResult.get("needSelectUnits"))) {
                result.put("success", true);
                result.put("needSelectUnits", true);
                result.put("units", matchResult.get("units"));
                result.put("fileName", new File(afterCopy).getName());
                result.put("message", "请选择单位后继续");
                return result;
            }
            afterMatchFile = new File((String) matchResult.get("filePath"));
            tempFile.delete();
            tempFile = afterMatchFile;

            // 第4步：处理扣款
            String afterDeduction = cwYzpzPmhService.processDeduction(
                    toMultipart(tempFile, new File((String) matchResult.get("filePath")).getName()),
                    dedFile, dedSheet, columnType, round, uploadDir);
            tempFile.delete();

            File finalFile = new File(afterDeduction);
            result.put("success", true);
            result.put("message", "核对完成（复制姓名、姓名匹配、处理扣款）");
            result.put("fileName", finalFile.getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "核对失败：" + e.getMessage());
            e.printStackTrace();
        } finally {
            // 清理临时文件
            if (afterCopyFile != null && afterCopyFile.exists()) afterCopyFile.delete();
            if (afterMatchFile != null && afterMatchFile.exists()) afterMatchFile.delete();
            if (tempFile != null && tempFile.exists()) tempFile.delete();
        }
        return result;
    }

    @PostMapping("/yzpz/save")
    public Map<String, Object> yzpzSave(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "outputDir", required = false) String outputDir) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String outputPath = cwYzpzPmhService.saveAs(file, outputDir != null ? outputDir : uploadDir);
            result.put("success", true);
            result.put("message", "保存完成");
            result.put("fileName", new File(outputPath).getName());
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "保存失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * 简单 MultipartFile 实现（用于链式调用）
     */
    private static class SimpleMultipartFile implements MultipartFile {
        private final byte[] content;
        private final String originalFilename;

        SimpleMultipartFile(File file, String originalFilename) throws IOException {
            this.content = java.nio.file.Files.readAllBytes(file.toPath());
            this.originalFilename = originalFilename;
        }

        @Override public String getName() { return "file"; }
        @Override public String getOriginalFilename() { return originalFilename; }
        @Override public String getContentType() { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }
        @Override public boolean isEmpty() { return content.length == 0; }
        @Override public long getSize() { return content.length; }
        @Override public byte[] getBytes() throws IOException { return content; }
        @Override public InputStream getInputStream() throws IOException { return new java.io.ByteArrayInputStream(content); }
        @Override public void transferTo(File dest) throws IOException, IllegalStateException { java.nio.file.Files.write(dest.toPath(), content); }
    }

    /**
     * 将 File 转换为 MultipartFile（用于链式调用）
     */
    private MultipartFile toMultipart(File file, String originalFilename) throws IOException {
        return new SimpleMultipartFile(file, originalFilename);
    }
}
