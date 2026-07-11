package com.example.project.controller.finance;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

/**
 * Finance 模块共享基类（不注册为 Controller，仅提供工具方法和共享字段）
 */
public abstract class FinanceBaseController {

    @Value("${file.upload-dir:./uploads}")
    protected String uploadDir;

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
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=\"" + file.getName() + "\"")
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
