package com.example.project.controller.finance.cwgjx;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cwgjx.CwgjxDgfCfService;
import com.example.project.task.MemoryMultipartFile;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/api/finance")
public class CwgjxDgfCfController extends FinanceBaseController {

    @Autowired
    private CwgjxDgfCfService cwgjxDgfCfService;

    /**
     * POST /api/finance/toolbox/labor-split
     * 功能：代工费拆分（cw_dgf_cf_xzy.py + dgf_cf_frame.py）
     * 异步化（双文件）：提交即返回 taskId，后台线程池执行
     */
    @PostMapping("/toolbox/labor-split")
    public Map<String, Object> toolboxLaborSplit(
            @RequestParam("file1") MultipartFile file1,
            @RequestParam("file2") MultipartFile file2,
            @RequestParam("sheet1Name") String sheet1Name,
            @RequestParam("sheet2Name") String sheet2Name,
            @RequestParam("sheetBName") String sheetBName,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        try {
            MemoryMultipartFile f1 = pin(file1);
            MemoryMultipartFile f2 = pin(file2);
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "toolbox/labor-split", "代工费拆分",
                    file1.getOriginalFilename(),
                    () -> {
                        File dir = new File(uploadDir);
                        if (!dir.exists()) dir.mkdirs();
                        CwgjxDgfCfService.ProcessResult pr = cwgjxDgfCfService.process(f1, sheet1Name, f2, sheet2Name, sheetBName, uploadDir);
                        return buildProcessResult(pr, "代工费拆分完成");
                    });
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件读取失败：" + e.getMessage());
            return err;
        }
    }

    /**
     * POST /api/finance/toolbox/labor-split-batch
     * 功能：代工费批量拆分（对齐 Python dgf_cf_frame.py 多文件 for 循环）
     * 上传多个代工费表格 + 1个发放记录表格，bState 跨文件累加处理
     * 异步化（多文件）：Service 签名需 File[]，请求线程内先落盘到 uploads 临时文件，
     * work 执行完清理临时文件；提交即返回 taskId
     */
    @PostMapping("/toolbox/labor-split-batch")
    public Map<String, Object> toolboxLaborSplitBatch(
            @RequestParam("file1") MultipartFile[] file1List,
            @RequestParam("file2") MultipartFile file2,
            @RequestParam("sheet1Name") String sheet1Names,
            @RequestParam("sheet2Name") String sheet2Name,
            @RequestParam("sheetBName") String sheetBName,
            @RequestHeader(value = "X-Device-Id", required = false) String deviceIdHeader) {
        File tmpDir = new File(uploadDir, "tmp_" + UUID.randomUUID().toString().replace("-", ""));
        try {
            if (file1List == null || file1List.length == 0 || file2 == null || file2.isEmpty()) {
                Map<String, Object> err = new HashMap<>();
                err.put("success", false);
                err.put("message", "请上传代工费文件与发放记录文件");
                return err;
            }
            if (!tmpDir.exists()) tmpDir.mkdirs();
            File[] savedFile1List = new File[file1List.length];
            for (int i = 0; i < file1List.length; i++) {
                savedFile1List[i] = saveToDisk(file1List[i], tmpDir);
            }
            File savedFile2 = saveToDisk(file2, tmpDir);
            String[] sheet1Arr = sheet1Names.split(",");
            if (sheet1Arr.length == 1 && file1List.length > 1) {
                String singleName = sheet1Arr[0];
                sheet1Arr = new String[file1List.length];
                Arrays.fill(sheet1Arr, singleName);
            }
            final String[] sheet1Final = sheet1Arr;
            if (sheet1Final.length != file1List.length) {
                Map<String, Object> err = new HashMap<>();
                err.put("success", false);
                err.put("message", "sheet1Name 数量(" + sheet1Final.length + ")与文件数量(" + file1List.length + ")不匹配");
                return err;
            }
            final File savedF2 = savedFile2;
            return taskService.submitAndRespond(deviceId(deviceIdHeader), "finance", "toolbox/labor-split-batch", "代工费批量拆分",
                    file1List[0].getOriginalFilename(),
                    () -> {
                        try {
                            File dir = new File(uploadDir);
                            if (!dir.exists()) dir.mkdirs();
                            CwgjxDgfCfService.ProcessResult pr = cwgjxDgfCfService.processBatch(
                                    savedFile1List, sheet1Final, savedF2, sheet2Name, sheetBName, uploadDir);
                            return buildProcessResult(pr, "代工费批量拆分完成，共处理 " + savedFile1List.length + " 个文件");
                        } finally {
                            deleteQuietly(tmpDir);   // 无论成败清理临时文件
                        }
                    });
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件保存失败：" + e.getMessage());
            return err;
        }
    }

    /** 把 MultipartFile 内容写到 tmpDir 下独立文件（避免线程池中操作已失效的 MultipartFile） */
    private File saveToDisk(MultipartFile mf, File tmpDir) throws IOException {
        String orig = mf.getOriginalFilename();
        String name = (orig == null || orig.isBlank()) ? UUID.randomUUID().toString() : orig;
        File dest = new File(tmpDir, UUID.randomUUID().toString().replace("-", "") + "_" + name);
        try (java.io.InputStream in = mf.getInputStream();
             java.io.OutputStream out = new java.io.FileOutputStream(dest)) {
            byte[] buf = new byte[8192];
            int n;
            while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
        }
        return dest;
    }

    /** 递归删除临时目录（忽略失败） */
    private void deleteQuietly(File dir) {
        if (dir == null || !dir.exists()) return;
        File[] fs = dir.listFiles();
        if (fs != null) {
            for (File f : fs) {
                if (f.isDirectory()) deleteQuietly(f);
                else f.delete();
            }
        }
        dir.delete();
    }

    private Map<String, Object> buildProcessResult(CwgjxDgfCfService.ProcessResult pr, String message) {
        Map<String, Object> result = new HashMap<>();
        result.put("message", message);
        result.put("fileName", new File(pr.outputPath).getName());
        result.put("summaryFileName", pr.summaryPath != null ? new File(pr.summaryPath).getName() : "");
        result.put("summaryFileNames", pr.summaryPaths != null ?
                pr.summaryPaths.stream().map(p -> new File(p).getName()).collect(Collectors.toList()) :
                new ArrayList<>());
        result.put("summaryCount", pr.summaryCount);
        result.put("splitCount", pr.splitCount);
        result.put("remark1", pr.remark1);
        result.put("companyPrefix", pr.companyPrefix);
        result.put("step4Details", pr.step4Details);
        result.put("step5Details", pr.step5Details);
        result.put("step5Result", pr.step5Result);
        return result;
    }
}
