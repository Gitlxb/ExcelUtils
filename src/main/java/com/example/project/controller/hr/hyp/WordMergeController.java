package com.example.project.controller.hr.hyp;

import com.example.project.service.hr.hyp.WordMergeService;
import com.example.project.controller.hr.HrBaseController;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.*;

@RestController
@RequestMapping("/api/hr")
public class WordMergeController extends HrBaseController {

    private static final Logger log = LoggerFactory.getLogger(WordMergeController.class);

    @Autowired
    private WordMergeService wordMergeService;

    private final ObjectMapper objectMapper = new ObjectMapper();

    /**
     * POST /api/hr/word-merge/extract
     */
    @PostMapping("/word-merge/extract")
    public Map<String, Object> extract(@RequestParam("files") MultipartFile[] files) {
        Map<String, Object> result = new HashMap<>();
        try {
            log.info("[word-merge/extract] 收到 {} 个标准文档", files == null ? 0 : files.length);
            List<MultipartFile> list = Arrays.asList(files);
            List<WordMergeService.Category> categories = wordMergeService.extractCategories(list);
            log.info("[word-merge/extract] 提取到 {} 个类别", categories.size());
            result.put("success", true);
            result.put("categories", categories);
            result.put("message", "共提取 " + categories.size() + " 个类别");
        } catch (Exception e) {
            log.error("[word-merge/extract] 失败：", e);
            result.put("success", false);
            result.put("message", "提取失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/word-merge/search
     */
    @PostMapping("/word-merge/search")
    public Map<String, Object> search(
            @RequestParam("categoriesJson") String categoriesJson,
            @RequestParam("names") List<String> names) {
        Map<String, Object> result = new HashMap<>();
        try {
            log.info("[word-merge/search] 收到 {} 个候选文件名，类别JSON长度={}",
                    names == null ? 0 : names.size(), categoriesJson == null ? 0 : categoriesJson.length());
            List<WordMergeService.Category> categories = objectMapper.readValue(
                    categoriesJson, new TypeReference<List<WordMergeService.Category>>() {});
            wordMergeService.searchCategories(categories, names);
            int total = 0;
            for (WordMergeService.Category c : categories) total += c.matchedFiles.size();
            log.info("[word-merge/search] 完成，匹配文件 {} 个", total);
            result.put("success", true);
            result.put("categories", categories);
            result.put("message", "搜索完成，共匹配 " + total + " 个文件");
        } catch (Exception e) {
            log.error("[word-merge/search] 失败：", e);
            result.put("success", false);
            result.put("message", "搜索失败：" + e.getMessage());
        }
        return result;
    }

    /**
     * POST /api/hr/word-merge
     */
    @PostMapping("/word-merge")
    public Map<String, Object> merge(
            @RequestParam("categoriesJson") String categoriesJson,
            @RequestParam("files") MultipartFile[] files) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();

            // ========== 详细诊断日志 ==========
            int fileCount = files == null ? 0 : files.length;
            log.info("========== [word-merge] START ==========");
            log.info("[word-merge] 收到 {} 个候选文件，categoriesJson 长度={}",
                    fileCount, categoriesJson == null ? 0 : categoriesJson.length());

            // 打印前 5 个文件名（含 originalFilename 和 name 对比）
            if (files != null) {
                for (int i = 0; i < Math.min(fileCount, 5); i++) {
                    MultipartFile f = files[i];
                    log.info("[word-merge] file[{}]: originalFilename='{}', name='{}', size={} bytes",
                            i, f.getOriginalFilename(), f.getName(), f.getSize());
                }
                if (fileCount > 5) log.info("[word-merge] ... 还有 {} 个文件省略", fileCount - 5);
            }

            List<WordMergeService.Category> categories = objectMapper.readValue(
                    categoriesJson, new TypeReference<List<WordMergeService.Category>>() {});

            // 打印所有 matchedFiles（关键诊断）
            log.info("[word-merge] 共 {} 个类别", categories.size());
            int catIdx = 0;
            for (WordMergeService.Category c : categories) {
                log.info("[word-merge]   类别[{}] parent='{}' matchedFiles={}",
                        catIdx++, c.parent, c.matchedFiles);
            }

            // 构建 candidateMap，webkitdirectory 下 originalFilename 含路径前缀（"章节/xx.docx"），
            // 须去掉前缀，与前端 File.name（纯 basename）和搜索结果的 matchedFiles 名对齐。
            Map<String, MultipartFile> candidateMap = new HashMap<>();
            for (MultipartFile f : files) {
                String orig = f.getOriginalFilename();
                if (orig != null) {
                    // 去掉 webkitdirectory 路径前缀（如 "章节/管理服务组织机构.docx" → "管理服务组织机构.docx"）
                    String basename = orig.contains("/") ? orig.substring(orig.lastIndexOf('/') + 1) : orig;
                    basename = basename.contains("\\") ? basename.substring(basename.lastIndexOf('\\') + 1) : basename;
                    candidateMap.put(basename, f);
                    if (!basename.equals(orig)) {
                        log.info("[word-merge] 修正文件名: '{}' → basename '{}'", orig, basename);
                    }
                }
            }
            log.info("[word-merge] candidateMap keys 数量={}，前10个: {}",
                    candidateMap.size(),
                    new ArrayList<>(candidateMap.keySet()).subList(0, Math.min(10, candidateMap.size())));

            // 交叉检查：每个分类第一个 matchedFile 是否在 candidateMap 中
            for (WordMergeService.Category c : categories) {
                if (!c.matchedFiles.isEmpty()) {
                    String first = c.matchedFiles.get(0).fileName;
                    boolean found = candidateMap.containsKey(first);
                    log.info("[word-merge]   分类 '{}' 第一个文件 '{}'（{}）在 candidateMap 中? {}",
                            c.parent, first, c.matchedFiles.get(0).matchType, found);
                    if (!found) {
                        // 打印相近的 key 提示
                        log.warn("[word-merge]   ⚠️ 缺失文件 '{}'，candidateMap 所有 keys: {}",
                                first, candidateMap.keySet());
                    }
                }
            }
            log.info("========== [word-merge] 调用 Service.mergeDocuments ==========");

            List<String> detailLogs = new ArrayList<>();
            String outputPath = wordMergeService.mergeDocuments(categories, candidateMap, uploadDir, detailLogs);
            log.info("[word-merge] ✓ 合并完成: {}", outputPath);
            result.put("success", true);
            result.put("message", "Word 文档合并完成");
            result.put("fileName", new File(outputPath).getName());
            result.put("logs", detailLogs);
        } catch (Exception e) {
            log.error("[word-merge] ✗ 合并失败，完整堆栈：", e);
            result.put("success", false);
            result.put("message", "合并失败：" + e.getClass().getSimpleName() + ": " + e.getMessage());
        }
        return result;
    }
}
