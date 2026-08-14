package com.example.project.controller.finance;

import com.example.project.service.finance.HbgsService;
import com.example.project.service.finance.HbgsService.MergeResult;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 合并个税 Controller
 * 对应 Python: hbgs_gui.py 的交互逻辑
 *
 * 端点：
 * - POST /api/finance/hbgs/preview-rules  预检规则表
 * - POST /api/finance/hbgs/merge          合并个税（上传规则表 + 文件夹数据文件）
 */
@RestController
@RequestMapping("/api/finance")
public class HbgsController extends FinanceBaseController {

    @Autowired
    private HbgsService hbgsService;

    /**
     * POST /api/finance/hbgs/preview-rules
     * 预检规则表：上传规则表文件，返回规则条数和表信息
     * 对应 Python hbgs_gui.py 中 _browse_rules() 选择规则表后立即加载验证的逻辑
     */
    @PostMapping("/hbgs/preview-rules")
    public Map<String, Object> previewRules(@RequestParam("rulesFile") MultipartFile rulesFile) {
        Map<String, Object> result = new HashMap<>();
        File tempFile = null;
        try {
            if (rulesFile == null || rulesFile.isEmpty()) {
                result.put("success", false);
                result.put("message", "请选择规则表文件");
                return result;
            }
            // 保存到临时文件
            tempFile = saveTempFile(rulesFile);
            // 加载规则信息
            List<String> logLines = new ArrayList<>();
            Map<String, Object[]> rules = hbgsService.loadFeeRules(tempFile, logLines::add);
            result.put("success", true);
            result.put("ruleCount", rules.size());
            result.put("worksheet", logLines.isEmpty() ? "" : logLines.get(0)); // 第一个日志行是工作表名
            result.put("logs", logLines);
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "读取规则表失败：" + e.getMessage());
        } finally {
            if (tempFile != null && tempFile.exists()) tempFile.delete();
        }
        return result;
    }

    /**
     * POST /api/finance/hbgs/merge
     * 合并个税：上传规则表 + 文件夹内所有 Excel 文件
     * 对应 Python hbgs_gui.py 中 _start() → _worker() → run_merge_collect() → write_merge_output() 的完整流程
     *
     * @param rulesFile 手续费规则表
     * @param dataFiles 文件夹中所有 Excel 文件（含子目录）
     */
    @PostMapping("/hbgs/merge")
    public Map<String, Object> merge(
            @RequestParam("rulesFile") MultipartFile rulesFile,
            @RequestParam("dataFiles") MultipartFile[] dataFiles) {
        Map<String, Object> result = new HashMap<>();

        File rulesTempFile = null;
        File dataTempDir = null;
        List<String> logLines = new ArrayList<>();
        StringBuilder logBuilder = new StringBuilder();

        try {
            // ==== 校验 ====
            if (rulesFile == null || rulesFile.isEmpty()) {
                result.put("success", false);
                result.put("message", "请上传手续费规则表");
                return result;
            }
            if (dataFiles == null || dataFiles.length == 0) {
                result.put("success", false);
                result.put("message", "请选择包含个税统计表的文件夹");
                return result;
            }

            // ==== 保存规则表到临时文件 ====
            rulesTempFile = saveTempFile(rulesFile);

            // ==== 保存数据文件到临时目录 ====
            dataTempDir = Files.createTempDirectory("hbgs_data_").toFile();
            List<File> savedDataFiles = new ArrayList<>();
            for (MultipartFile mf : dataFiles) {
                if (mf.isEmpty()) continue;
                String fn = mf.getOriginalFilename();
                if (fn == null) fn = "unknown.xlsx";
                // 只处理 Excel 文件
                String lower = fn.toLowerCase();
                if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) continue;
                if (fn.startsWith("~$")) continue;
                File dest = new File(dataTempDir, fn);
                // webkitdirectory 上传时文件名含子目录路径，需确保父目录存在
                File parentDir = dest.getParentFile();
                if (parentDir != null && !parentDir.exists()) {
                    parentDir.mkdirs();
                }
                // 处理重名
                if (dest.exists()) {
                    String base = fn.substring(0, fn.lastIndexOf('.'));
                    String ext = fn.substring(fn.lastIndexOf('.'));
                    dest = new File(dataTempDir, base + "_" + System.currentTimeMillis() + ext);
                }
                mf.transferTo(dest);
                savedDataFiles.add(dest);
            }

            if (savedDataFiles.isEmpty()) {
                result.put("success", false);
                result.put("message", "文件夹中没有找到有效的 Excel 文件（.xlsx/.xls）");
                return result;
            }

            // ==== 执行合并逻辑（对应 Python _worker → run_merge_collect） ====
            logLines.add("开始合并...");
            MergeResult mergeResult = hbgsService.runMergeCollect(savedDataFiles, rulesTempFile,
                    msg -> {
                        synchronized (logLines) {
                            logLines.add(msg);
                            logBuilder.append(msg).append("\n");
                        }
                    });

            if (mergeResult.header == null || mergeResult.rows.isEmpty()) {
                result.put("success", false);
                result.put("message", "未能生成汇总数据，请查看日志");
                result.put("logs", logLines);
                result.put("totalFiles", mergeResult.totalFiles);
                result.put("skipped", mergeResult.skipped);
                result.put("errors", mergeResult.errors);
                return result;
            }

            // ==== 写出输出文件（对应 Python _on_done → write_merge_output） ====
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
            String outputFileName = "个税汇总" + timestamp + ".xlsx";
            File outputFile = new File(dir, outputFileName);
            hbgsService.writeMergeOutput(outputFile, mergeResult.header, mergeResult.rows, mergeResult.targetIdCol);

            // ==== 组装返回结果 ====
            result.put("success", true);
            result.put("message", "合并完成");
            result.put("fileName", outputFileName);
            result.put("totalFiles", mergeResult.totalFiles);
            result.put("totalRows", mergeResult.totalRows);
            result.put("skipped", mergeResult.skipped);
            result.put("errors", mergeResult.errors);
            result.put("logs", logLines);

            // 两类提示（对应 Python 末尾 MsgBox 的 ⚠ 文字）
            List<String> warnings = new ArrayList<>();
            if (!mergeResult.emptyKeywordOps.isEmpty()) {
                warnings.add("规则表有 " + mergeResult.emptyKeywordOps.size() + " 个运营项目未填写扣费关键词");
            }
            if (!mergeResult.missingOps.isEmpty()) {
                warnings.add("源文件有 " + mergeResult.missingOps.size() + " 个运营项目缺规则");
            }
            if (!warnings.isEmpty()) {
                result.put("warnings", warnings);
                result.put("emptyKeywordOps", new ArrayList<>(mergeResult.emptyKeywordOps));
                result.put("missingOps", new ArrayList<>(mergeResult.missingOps));
            }

        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "处理失败：" + e.getMessage());
            result.put("logs", logLines);
        } finally {
            // 清理临时文件
            if (rulesTempFile != null && rulesTempFile.exists()) rulesTempFile.delete();
            if (dataTempDir != null && dataTempDir.exists()) deleteDirectory(dataTempDir);
        }

        return result;
    }

    // ==================== 辅助方法 ====================

    /** 保存 MultipartFile 到临时文件 */
    private File saveTempFile(MultipartFile mf) throws IOException {
        String originalName = mf.getOriginalFilename();
        String suffix = ".tmp";
        if (originalName != null) {
            int dot = originalName.lastIndexOf('.');
            if (dot > 0) suffix = originalName.substring(dot);
        }
        File tempFile = File.createTempFile("hbgs_rules_", suffix);
        mf.transferTo(tempFile);
        return tempFile;
    }

    /** 递归删除目录 */
    private void deleteDirectory(File dir) {
        if (dir == null || !dir.exists()) return;
        File[] files = dir.listFiles();
        if (files != null) {
            for (File f : files) {
                if (f.isDirectory()) deleteDirectory(f);
                else f.delete();
            }
        }
        dir.delete();
    }
}
