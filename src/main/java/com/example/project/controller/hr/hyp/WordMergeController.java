package com.example.project.controller.hr.hyp;

import com.example.project.service.hr.hyp.WordMergeService;
import com.example.project.controller.hr.HrBaseController;
import com.example.project.task.MemoryMultipartFile;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
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
     * Word 文档合并（多文件 + 分类状态，异步化：提交即返回 taskId，后台执行）
     */
    @PostMapping("/word-merge")
    public Map<String, Object> merge(
            @RequestParam("categoriesJson") String categoriesJson,
            @RequestParam("files") MultipartFile[] files,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            if (files == null || files.length == 0) {
                Map<String, Object> err = new HashMap<>();
                err.put("success", false);
                err.put("message", "请选择候选文件");
                return err;
            }
            // 请求线程内解析分类状态 + 固化所有文件为内存副本（异步线程安全）
            List<WordMergeService.Category> categories = objectMapper.readValue(
                    categoriesJson, new TypeReference<List<WordMergeService.Category>>() {});
            Map<String, MultipartFile> candidateMap = new HashMap<>();
            for (MultipartFile f : files) {
                if (f == null || f.isEmpty()) continue;
                String orig = f.getOriginalFilename();
                if (orig == null) continue;
                String basename = orig.contains("/") ? orig.substring(orig.lastIndexOf('/') + 1) : orig;
                basename = basename.contains("\\") ? basename.substring(basename.lastIndexOf('\\') + 1) : basename;
                candidateMap.put(basename, pin(f));
            }
            if (candidateMap.isEmpty()) {
                Map<String, Object> err = new HashMap<>();
                err.put("success", false);
                err.put("message", "没有有效的候选文件");
                return err;
            }
            String display = "Word文档合并";
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "hr", "word-merge", display,
                    null, () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        List<String> detailLogs = new ArrayList<>();
                        String outputPath = wordMergeService.mergeDocuments(categories, candidateMap, uploadDir, detailLogs);
                        Map<String, Object> result = new HashMap<>();
                        result.put("message", "Word 文档合并完成");
                        result.put("fileName", new File(outputPath).getName());
                        result.put("logs", detailLogs);
                        return result;
                    });
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件读取失败：" + e.getMessage());
            return err;
        } catch (Exception e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "请求解析失败：" + e.getClass().getSimpleName() + ": " + e.getMessage());
            return err;
        }
    }
}
