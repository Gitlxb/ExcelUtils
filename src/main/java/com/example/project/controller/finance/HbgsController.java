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

    /** 规则表固定文件名（2026-10-26：不再单独上传，改为从所选目录内自动识别） */
    private static final String RULES_FILE_BASE = "收取手续费银行信息";

    /**
     * POST /api/finance/hbgs/merge
     * 合并个税：只需选择数据文件夹，规则表从目录内自动识别
     *
     * 规则表来源（唯一，无兜底）：
     *   本次上传的文件中必须存在「收取手续费银行信息.xlsx」（或 .xls），
     *   否则直接返回失败并提示，不再产出「手续费整列为空」的结果。
     *
     * 异步化（重处理，可能数十文件）：请求线程内把规则表 + 数据文件落盘到临时目录，
     * 提交 taskId 后后台线程池执行合并；logs / warnings / 统计字段均写入 task.result，
     * 前端轮询 /api/tasks/{taskId} 取结果。
     *
     * @param dataFiles 文件夹中所有 Excel 文件（含子目录，其中须包含规则表）
     */
    @PostMapping("/hbgs/merge")
    public Map<String, Object> merge(
            @RequestParam("dataFiles") MultipartFile[] dataFiles,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        File rulesTempFile = null;
        File dataTempDir = null;
        try {
            // ==== 校验 ====
            if (dataFiles == null || dataFiles.length == 0) {
                return fail("请选择包含个税统计表的文件夹");
            }

            // ==== ① 目录内识别规则表（唯一来源，找不到即停止）====
            List<MultipartFile> rulesMatches = findRulesFiles(dataFiles);
            if (rulesMatches.isEmpty()) {
                return fail("所选择的目录需要包含「" + RULES_FILE_BASE + ".xlsx」，未检测到该文件，已停止处理。");
            }
            final MultipartFile rulesMf = rulesMatches.get(0);

            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // ==== ② 落盘：规则表单独存，且不计入数据文件 ====
            rulesTempFile = saveTempFile(rulesMf);
            dataTempDir = java.nio.file.Files.createTempDirectory("hbgs_data_").toFile();
            List<File> savedDataFiles = new ArrayList<>();
            int seq = 0;
            for (MultipartFile mf : dataFiles) {
                if (mf == null || mf.isEmpty()) continue;
                // 规则表不进数据清单：否则会被当成"无个税统计表"的文件而计入 errors/skipped
                if (rulesMatches.contains(mf)) continue;
                // 去掉路径前缀：防目录穿越，也保证汇总表「文件名」列干净
                String fn = baseName(mf.getOriginalFilename());
                if (fn.isEmpty()) fn = "unknown.xlsx";
                String lower = fn.toLowerCase();
                if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) continue;
                if (fn.startsWith("~$")) continue;
                // 每个文件独占一个子目录：不同子目录的同名文件互不覆盖。
                // 旧实现遇到同名会改名 base_时间戳.ext，而该名字会被写进汇总表「文件名」列（数据污染）。
                File slot = new File(dataTempDir, String.format("%04d", ++seq));
                if (!slot.exists() && !slot.mkdirs()) {
                    throw new IOException("无法创建临时目录：" + slot.getAbsolutePath());
                }
                File dest = new File(slot, fn);
                mf.transferTo(dest);
                savedDataFiles.add(dest);
            }
            if (savedDataFiles.isEmpty()) {
                cleanup(rulesTempFile, dataTempDir);
                return fail("目录中除规则表外没有其他 Excel 文件（.xlsx/.xls）");
            }

            final File rulesF = rulesTempFile;
            final File dataD = dataTempDir;
            final List<File> dataFs = savedDataFiles;
            final String rulesName = rulesMf.getOriginalFilename();
            final int rulesFound = rulesMatches.size();
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "hbgs/merge", "合并个税",
                    rulesName,
                    () -> {
                        List<String> logLines = new ArrayList<>();
                        StringBuilder logBuilder = new StringBuilder();
                        try {
                            logLines.add("开始合并...");
                            logLines.add("规则表：目录内识别 -> " + rulesName
                                    + (rulesFound > 1 ? "（目录中有 " + rulesFound + " 个同名文件，已使用第一个）" : ""));
                            MergeResult mergeResult = hbgsService.runMergeCollect(dataFs, rulesF,
                                    msg -> {
                                        synchronized (logLines) {
                                            logLines.add(msg);
                                            logBuilder.append(msg).append("\n");
                                        }
                                    });

                            // 规则表解析不出任何规则 → 明确失败，不产出"手续费整列留空"的表
                            if (mergeResult.ruleCount == 0) {
                                Map<String, Object> r = new HashMap<>();
                                r.put("success", false);
                                r.put("message", "规则表「" + rulesName + "」未解析出任何规则：请检查「运营项目」「扣费银行（多填）」两列是否存在，以及「客户状态」是否全部为「终止合作」。");
                                r.put("logs", logLines);
                                return r;
                            }

                            if (mergeResult.header == null || mergeResult.rows.isEmpty()) {
                                Map<String, Object> r = new HashMap<>();
                                r.put("success", false);
                                r.put("message", "未能生成汇总数据，请查看日志");
                                r.put("logs", logLines);
                                r.put("totalFiles", mergeResult.totalFiles);
                                r.put("skipped", mergeResult.skipped);
                                r.put("errors", mergeResult.errors);
                                return r;   // TaskService 据 success=false 置 FAILED 并保留 result
                            }

                            String timestamp = new java.text.SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
                            String outputFileName = "个税汇总" + timestamp + ".xlsx";
                            File outputFile = new File(dir, outputFileName);
                            hbgsService.writeMergeOutput(outputFile, mergeResult.header, mergeResult.rows, mergeResult.targetIdCol);

                            Map<String, Object> r = new HashMap<>();
                            r.put("message", "合并完成");
                            r.put("fileName", outputFileName);
                            r.put("rulesFileName", rulesName);
                            r.put("ruleCount", mergeResult.ruleCount);
                            r.put("totalFiles", mergeResult.totalFiles);
                            r.put("totalRows", mergeResult.totalRows);
                            r.put("skipped", mergeResult.skipped);
                            r.put("errors", mergeResult.errors);
                            r.put("logs", logLines);

                            List<String> warnings = new ArrayList<>();
                            if (!mergeResult.emptyKeywordOps.isEmpty()) {
                                warnings.add("规则表有 " + mergeResult.emptyKeywordOps.size() + " 个运营项目未填写扣费关键词");
                            }
                            if (!mergeResult.missingOps.isEmpty()) {
                                warnings.add("源文件有 " + mergeResult.missingOps.size() + " 个运营项目缺规则");
                            }
                            if (!warnings.isEmpty()) {
                                r.put("warnings", warnings);
                                r.put("emptyKeywordOps", new ArrayList<>(mergeResult.emptyKeywordOps));
                                r.put("missingOps", new ArrayList<>(mergeResult.missingOps));
                            }
                            return r;
                        } finally {
                            if (rulesF != null && rulesF.exists()) rulesF.delete();
                            if (dataD != null && dataD.exists()) deleteDirectory(dataD);
                        }
                    });
        } catch (Exception e) {
            cleanup(rulesTempFile, dataTempDir);
            return fail("处理失败：" + e.getMessage());
        }
    }

    /** 统一失败返回 */
    private Map<String, Object> fail(String message) {
        Map<String, Object> err = new HashMap<>();
        err.put("success", false);
        err.put("message", message);
        return err;
    }

    /** 清理落盘临时文件/目录 */
    private void cleanup(File rulesTempFile, File dataTempDir) {
        if (rulesTempFile != null && rulesTempFile.exists()) rulesTempFile.delete();
        if (dataTempDir != null && dataTempDir.exists()) deleteDirectory(dataTempDir);
    }

    /** 取文件名（去掉任何路径前缀，防目录穿越并保证「文件名」列干净） */
    private String baseName(String name) {
        if (name == null) return "";
        int slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
        return slash >= 0 ? name.substring(slash + 1) : name;
    }

    /**
     * 在本次上传的文件中定位规则表：文件名（去扩展名）等于「收取手续费银行信息」，扩展名 .xlsx / .xls
     *
     * 说明：浏览器只回传文件名（不含目录层级），因此只能按文件名匹配；
     *       目录内若存在多份同名文件（不同子目录），取第一个并记入日志。
     */
    private List<MultipartFile> findRulesFiles(MultipartFile[] files) {
        List<MultipartFile> matched = new ArrayList<>();
        if (files == null) return matched;
        for (MultipartFile mf : files) {
            if (mf == null || mf.isEmpty()) continue;
            String fn = baseName(mf.getOriginalFilename());
            if (fn.isEmpty()) continue;
            int dot = fn.lastIndexOf('.');
            if (dot <= 0) continue;
            String base = fn.substring(0, dot).trim();
            String ext = fn.substring(dot + 1).toLowerCase();
            if (!"xlsx".equals(ext) && !"xls".equals(ext)) continue;
            if (RULES_FILE_BASE.equals(base)) matched.add(mf);
        }
        return matched;
    }

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
