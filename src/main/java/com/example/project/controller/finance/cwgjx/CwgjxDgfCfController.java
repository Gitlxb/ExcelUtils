package com.example.project.controller.finance.cwgjx;

import com.example.project.controller.finance.FinanceBaseController;
import com.example.project.service.finance.cwgjx.CwgjxDgfCfService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
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
     */
    @PostMapping("/toolbox/labor-split")
    public Map<String, Object> toolboxLaborSplit(
            @RequestParam("file1") MultipartFile file1,
            @RequestParam("file2") MultipartFile file2,
            @RequestParam("sheet1Name") String sheet1Name,
            @RequestParam("sheet2Name") String sheet2Name,
            @RequestParam("sheetBName") String sheetBName) {
        Map<String, Object> result = new HashMap<>();
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            CwgjxDgfCfService.ProcessResult processResult = cwgjxDgfCfService.process(file1, sheet1Name, file2, sheet2Name, sheetBName, uploadDir);
            result.put("success", true);
            result.put("message", "代工费拆分完成");
            result.put("fileName", new File(processResult.outputPath).getName());
            result.put("summaryFileName", new File(processResult.summaryPath).getName());
            result.put("summaryCount", processResult.summaryCount);
            result.put("splitCount", processResult.splitCount);
            result.put("remark1", processResult.remark1);
            result.put("companyPrefix", processResult.companyPrefix);
            result.put("step4Details", processResult.step4Details);
            result.put("step5Details", processResult.step5Details);
            result.put("step5Result", processResult.step5Result);
        } catch (Throwable t) {
            // 捕获所有异常，包括 Error，确保前端一定能收到 JSON 响应
            result.put("success", false);
            String msg = t.getMessage();
            if (msg == null || msg.isEmpty()) {
                msg = t.getClass().getSimpleName();
            }
            result.put("message", "处理失败：" + msg);
            t.printStackTrace(); // 输出完整堆栈到后端日志
        }
        return result;
    }

    /**
     * POST /api/finance/toolbox/labor-split-batch
     * 功能：代工费批量拆分（对齐 Python dgf_cf_frame.py 多文件 for 循环）
     * 上传多个代工费表格 + 1个发放记录表格，bState 跨文件累加处理
     */
    @PostMapping("/toolbox/labor-split-batch")
    public Map<String, Object> toolboxLaborSplitBatch(
            @RequestParam("file1") MultipartFile[] file1List,
            @RequestParam("file2") MultipartFile file2,
            @RequestParam("sheet1Name") String sheet1Names,
            @RequestParam("sheet2Name") String sheet2Name,
            @RequestParam("sheetBName") String sheetBName) {
        Map<String, Object> result = new HashMap<>();
        // 保存所有 files1 的临时文件路径，用于 finally 清理
        List<File> savedFiles = new ArrayList<>();
        File savedFile2 = null;
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            String[] sheet1Arr = sheet1Names.split(",");
            // 如果只传了一个工作表名，所有文件共用同一张工作表
            if (sheet1Arr.length == 1 && file1List.length > 1) {
                String singleName = sheet1Arr[0];
                sheet1Arr = new String[file1List.length];
                Arrays.fill(sheet1Arr, singleName);
            }
            if (sheet1Arr.length != file1List.length) {
                result.put("success", false);
                result.put("message", "sheet1Name 数量(" + sheet1Arr.length + ")与文件数量(" + file1List.length + ")不匹配");
                return result;
            }
            // 立即保存所有 MultipartFile 到磁盘（流拷贝，避免 transferTo 的 Temp 依赖）
            File[] savedFile1List = new File[file1List.length];
            for (int i = 0; i < file1List.length; i++) {
                File saved = new File(dir, UUID.randomUUID() + "_" + file1List[i].getOriginalFilename());
                try (InputStream in = file1List[i].getInputStream();
                     OutputStream out = new java.io.FileOutputStream(saved)) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
                }
                savedFile1List[i] = saved;
                savedFiles.add(saved);
            }
            savedFile2 = new File(dir, UUID.randomUUID() + "_" + file2.getOriginalFilename());
            try (InputStream in = file2.getInputStream();
                 OutputStream out = new java.io.FileOutputStream(savedFile2)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            System.out.println("[CwgjxDgfCfController] 批量文件已保存到磁盘，共 " + file1List.length + " 个代工费文件 + 1 个发放记录文件");

            CwgjxDgfCfService.ProcessResult processResult = cwgjxDgfCfService.processBatch(
                    savedFile1List, sheet1Arr, savedFile2, sheet2Name, sheetBName, uploadDir);
            result.put("success", true);
            result.put("message", "代工费批量拆分完成，共处理 " + file1List.length + " 个文件");
            result.put("fileName", new File(processResult.outputPath).getName());
            result.put("summaryFileName", processResult.summaryPath != null ? new File(processResult.summaryPath).getName() : "");
            result.put("summaryFileNames", processResult.summaryPaths != null ?
                    processResult.summaryPaths.stream().map(p -> new File(p).getName()).collect(Collectors.toList()) :
                    new ArrayList<>());
            result.put("summaryCount", processResult.summaryCount);
            result.put("splitCount", processResult.splitCount);
            result.put("remark1", processResult.remark1);
            result.put("companyPrefix", processResult.companyPrefix);
            result.put("step4Details", processResult.step4Details);
            result.put("step5Details", processResult.step5Details);
            result.put("step5Result", processResult.step5Result);
        } catch (Throwable t) {
            result.put("success", false);
            String msg = t.getMessage();
            if (msg == null || msg.isEmpty()) msg = t.getClass().getSimpleName();
            result.put("message", "处理失败：" + msg);
            t.printStackTrace();
        } finally {
            // 清理 Controller 层保存的临时文件（Service 内部已处理自己的临时文件）
            for (File f : savedFiles) {
                if (f != null && f.exists()) {
                    f.delete();
                }
            }
        }
        return result;
    }
}
