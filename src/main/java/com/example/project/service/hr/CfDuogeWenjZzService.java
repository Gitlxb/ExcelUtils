package com.example.project.service.hr;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 对应 Python: cf_duoge_wenj_zz.py
 * 功能：按列拆分 Excel 为多个独立文件，支持二次拆分为工作表
 * 特点：长数字文本保护、日期格式化、文件名清理、工作表排序
 */
@Service
public class CfDuogeWenjZzService {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    /**
     * 按列拆分 Excel 为多个独立文件，打包为 ZIP
     * @param file        上传的 Excel 文件
     * @param colLetters  要拆分的列字母（多个用逗号分隔）
     * @param outputDir   输出目录
     * @return ZIP 文件路径
     */
    public String splitToFiles(MultipartFile file, String colLetters, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            file.transferTo(inputFile.toPath());
            File zipFile = new File(outputDir, "split_hr_" + System.currentTimeMillis() + ".zip");
            processSplit(inputFile, zipFile, colLetters);
            return zipFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    private void processSplit(File inputFile, File zipFile, String colLetters) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            if (lastRow < 1) throw new Exception("Excel 文件没有数据");

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new Exception("Excel 文件没有表头行");

            // 解析列字母
            String[] letters = colLetters.split(",");
            int[] colIndices = new int[letters.length];
            for (int i = 0; i < letters.length; i++) {
                colIndices[i] = ExcelUtils.excelColumnToIndex(letters[i].trim());
            }

            // 计算最大列数
            int maxCols = 0;
            for (int r = 0; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row != null && row.getLastCellNum() > maxCols) {
                    maxCols = row.getLastCellNum();
                }
            }
            int colCount = Math.max(maxCols, 1);

            // 收集所有数据行
            List<Integer> allRows = new ArrayList<>();
            for (int r = 1; r <= lastRow; r++) {
                if (sheet.getRow(r) != null) allRows.add(r);
            }

            // 按指定列的值分组（多列组合作为 key）
            Map<String, List<Integer>> groups = new LinkedHashMap<>();
            for (int rowIndex : allRows) {
                Row row = sheet.getRow(rowIndex);
                StringBuilder key = new StringBuilder();
                for (int ci : colIndices) {
                    String val = ExcelUtils.getCellStringValue(row.getCell(ci));
                    if (key.length() > 0) key.append("_");
                    key.append(val != null && !val.isEmpty() ? val : "空值");
                }
                groups.computeIfAbsent(key.toString(), k -> new ArrayList<>()).add(rowIndex);
            }

            // 按字母倒序排序 key
            List<String> sortedKeys = new ArrayList<>(groups.keySet());
            sortedKeys.sort((a, b) -> b.compareToIgnoreCase(a));

            // 写入 ZIP
            try (FileOutputStream fos = new FileOutputStream(zipFile);
                 ZipOutputStream zos = new ZipOutputStream(fos)) {

                for (String key : sortedKeys) {
                    List<Integer> rowIndices = groups.get(key);

                    // 生成文件名（清理非法字符 + 日期格式化）
                    String fileName = formatFileName(key);
                    fileName = fileName.replaceAll("[<>:\"/\\\\|?*]", "_");
                    if (fileName.length() > 100) fileName = fileName.substring(0, 100);
                    fileName += ".xlsx";

                    try (Workbook splitWb = new XSSFWorkbook()) {
                        Sheet splitSheet = splitWb.createSheet("Sheet1");

                        // 复制表头
                        Row newHeader = splitSheet.createRow(0);
                        ExcelUtils.copyRow(headerRow, newHeader, 0, colCount - 1);

                        // 复制数据行（所有值按文本处理）
                        int newIdx = 1;
                        for (int rowIndex : rowIndices) {
                            Row srcRow = sheet.getRow(rowIndex);
                            if (srcRow == null) continue;
                            Row destRow = splitSheet.createRow(newIdx++);
                            for (int c = 0; c < colCount; c++) {
                                Cell srcCell = srcRow.getCell(c);
                                Cell destCell = destRow.createCell(c);
                                if (srcCell != null) {
                                    String val = ExcelUtils.getCellStringValue(srcCell);
                                    // 检测是否为日期格式，统一为 YYYY-MM-DD
                                    val = formatDateString(val);
                                    destCell.setCellValue(val);
                                    // 长数字（10位以上）设为文本格式
                                    if (val != null && val.matches("\\d{11,}")) {
                                        destCell.setCellType(CellType.STRING);
                                    }
                                }
                            }
                        }

                        zos.putNextEntry(new ZipEntry(fileName));
                        splitWb.write(zos);
                        zos.closeEntry();
                    }
                }
            }
        }
    }

    /**
     * 格式化值为文件名友好格式（日期转为 YYYY-MM-DD）
     */
    private String formatFileName(String value) {
        if (value == null || value.isEmpty()) return "空值";
        String dateStr = formatDateString(value);
        return dateStr != null ? dateStr : value;
    }

    /**
     * 尝试将字符串解析为日期并格式化为 YYYY-MM-DD
     * 如果不是日期格式，返回原字符串
     */
    private String formatDateString(String value) {
        if (value == null || value.isEmpty()) return value;

        DateTimeFormatter[] formatters = {
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("yyyyMMdd"),
            DateTimeFormatter.ofPattern("yyyy年MM月dd日")
        };

        // 先尝试直接解析
        for (DateTimeFormatter fmt : formatters) {
            try {
                LocalDate date = LocalDate.parse(value, fmt);
                return date.format(DATE_FMT);
            } catch (DateTimeParseException ignored) {}
        }

        // 尝试取空格前的日期部分（去除时间）
        if (value.contains(" ")) {
            String datePart = value.split(" ")[0];
            for (DateTimeFormatter fmt : formatters) {
                try {
                    LocalDate date = LocalDate.parse(datePart, fmt);
                    return date.format(DATE_FMT);
                } catch (DateTimeParseException ignored) {}
            }
        }

        return value; // 不是日期，原样返回
    }
}
