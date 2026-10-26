package com.example.project.controller.finance;

import com.example.project.task.FileTaskWork;
import com.example.project.task.MemoryMultipartFile;
import com.example.project.task.TaskService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Finance 模块共享基类（不注册为 Controller，仅提供工具方法和共享字段）
 */
public abstract class FinanceBaseController {

    @Value("${file.upload-dir:./uploads}")
    protected String uploadDir;

    @Autowired
    protected TaskService taskService;

    /**
     * 异步提交便捷方法（单文件型工具统一入口）：请求线程内固化文件为内存副本 →
     * 提交后台线程池 → 立即返回 {success,taskId}。work 抛异常由 TaskService 置 FAILED。
     */
    protected Map<String, Object> doSubmitAsync(MultipartFile file, String deviceId,
                                                String module, String apiName,
                                                String displayName, FileTaskWork work) {
        String originalName = file == null ? null : file.getOriginalFilename();
        try {
            byte[] bytes = file.getBytes();
            MemoryMultipartFile safe = new MemoryMultipartFile(
                    file.getName(), file.getOriginalFilename(), file.getContentType(), bytes);
            return taskService.submitAndRespond(deviceId, module, apiName, displayName,
                    originalName, () -> work.execute(safe));
        } catch (IOException e) {
            Map<String, Object> err = new HashMap<>();
            err.put("success", false);
            err.put("message", "文件读取失败：" + e.getMessage());
            return err;
        }
    }

    /** 取设备号（前端 X-Device-Id，缺失兼容单机/旧前端） */
    protected String deviceId(String header) {
        return (header == null || header.isBlank()) ? "" : header;
    }

    /** 把 MultipartFile 固化为内存副本（异步线程安全）：请求线程内调用，异步线程可安全 getBytes/getInputStream */
    protected MemoryMultipartFile pin(MultipartFile f) throws IOException {
        return new MemoryMultipartFile(f.getName(), f.getOriginalFilename(), f.getContentType(), f.getBytes());
    }

    // ==================== 共享工具方法（不注册端点，由 FinanceUtilController 暴露） ====================

    /**
     * 下载文件（供 UtilController 暴露为 /download/{fileName} 端点）
     */
    protected ResponseEntity<Resource> doDownloadFile(String fileName) {
        try {
            File file = new File(uploadDir, fileName);
            if (!file.exists()) {
                return ResponseEntity.notFound().build();
            }
            Resource resource = new FileSystemResource(file);
            // RFC 5987：中文文件名需 UTF-8 百分号编码，否则 Tomcat 判定 Content-Disposition 非法直接移除该头
            String encoded = URLEncoder.encode(file.getName(), StandardCharsets.UTF_8).replace("+", "%20");
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=\"" + encoded + "\"; filename*=UTF-8''" + encoded)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(resource);
        } catch (Exception e) {
            return ResponseEntity.internalServerError().build();
        }
    }

    /**
     * 读取工作表名称（供 UtilController 暴露为 /toolbox/sheets 端点）
     */
    protected Map<String, Object> doToolboxSheets(MultipartFile file) {
        return doReadSheets(file);
    }

    /**
     * 读取工作表名称（供 UtilController 暴露为 /summary/sheets 端点）
     */
    protected Map<String, Object> doSummarySheets(MultipartFile file) {
        return doReadSheets(file);
    }

    private Map<String, Object> doReadSheets(MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        File temp = null;
        try {
            File dir = new File(uploadDir);
            if (!dir.exists()) dir.mkdirs();
            temp = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
            try (InputStream in = file.getInputStream();
                 OutputStream out = new java.io.FileOutputStream(temp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            try (Workbook wb = WorkbookFactory.create(new FileInputStream(temp))) {
                List<String> sheets = new ArrayList<>();
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    sheets.add(wb.getSheetName(i));
                }
                result.put("success", true);
                result.put("sheets", sheets);
            }
        } catch (Exception e) {
            result.put("success", false);
            result.put("message", "读取工作表失败：" + e.getMessage());
        } finally {
            if (temp != null && temp.exists()) temp.delete();
        }
        return result;
    }

    // ==================== 共享预览方法 ====================

    /**
     * 读取 Excel 文件前 limit 行作为预览数据
     */
    protected Map<String, Object> readExcelPreview(File file, int limit) {
        Map<String, Object> preview = new HashMap<>();
        List<String> columns = new ArrayList<>();
        List<Map<String, Object>> rows = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    Cell cell = headerRow.getCell(c);
                    columns.add(cell != null ? cell.toString().trim() : "");
                }
            }
            int endRow = Math.min(limit + 1, sheet.getLastRowNum() + 1);
            for (int r = 1; r < endRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Map<String, Object> rowMap = new LinkedHashMap<>();
                for (int c = 0; c < columns.size(); c++) {
                    Cell cell = row.getCell(c);
                    String colName = columns.get(c);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                rowMap.put(colName, cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    rowMap.put(colName, cell.getDateCellValue());
                                } else {
                                    double d = cell.getNumericCellValue();
                                    rowMap.put(colName, d == Math.rint(d) ? (long) d : d);
                                }
                                break;
                            case BOOLEAN:
                                rowMap.put(colName, cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                CellType rt = cell.getCachedFormulaResultType();
                                if (rt == CellType.NUMERIC) {
                                    double d = cell.getNumericCellValue();
                                    rowMap.put(colName, d == Math.rint(d) ? (long) d : d);
                                } else if (rt == CellType.STRING) {
                                    rowMap.put(colName, cell.getStringCellValue());
                                } else {
                                    rowMap.put(colName, cell.toString());
                                }
                                break;
                            default:
                                rowMap.put(colName, cell.toString());
                                break;
                        }
                    } else {
                        rowMap.put(colName, "");
                    }
                }
                rows.add(rowMap);
            }
            preview.put("columns", columns);
            preview.put("data", rows);
        } catch (Exception e) {
            preview.put("columns", Collections.emptyList());
            preview.put("data", Collections.emptyList());
            preview.put("error", e.getMessage());
        }
        return preview;
    }
}
