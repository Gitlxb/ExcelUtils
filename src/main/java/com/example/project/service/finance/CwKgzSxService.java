package com.example.project.service.finance;

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

/**
 * 对应 Python: cw_kgz_sx.py
 * 功能：按日期范围筛选 Excel 数据
 * 规则：20号及之前 → 前3个月起；20号之后 → 前2个月起
 */
@Service
public class CwKgzSxService {

    /**
     * 按日期范围筛选数据
     * @param file           上传的 Excel 文件
     * @param dateColumnName 日期列的列名（表头）
     * @param outputDir      输出目录
     * @return 处理后的文件路径
     */
    public String filterByDate(MultipartFile file, String dateColumnName, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }
        if (dateColumnName == null || dateColumnName.trim().isEmpty()) {
            throw new IllegalArgumentException("请指定日期列名");
        }

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            file.transferTo(inputFile.toPath());
            File outputFile = new File(outputDir, "filtered_" + System.currentTimeMillis() + ".xlsx");
            processDateFilter(inputFile, outputFile, dateColumnName.trim());
            return outputFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    private void processDateFilter(File inputFile, File outputFile, String dateColumnName) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook srcWorkbook = new XSSFWorkbook(fis)) {

            Sheet srcSheet = srcWorkbook.getSheetAt(0);
            int lastRow = srcSheet.getLastRowNum();
            if (lastRow < 1) {
                throw new Exception("Excel 文件没有数据");
            }

            // 查找表头确定日期列索引
            Row headerRow = srcSheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("Excel 文件没有表头行");
            }

            int dateColIndex = -1;
            int maxCols = 0;
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                String header = ExcelUtils.getCellStringValue(headerRow.getCell(c));
                if (header != null && !header.isEmpty()) {
                    if (header.equals(dateColumnName)) {
                        dateColIndex = c;
                    }
                    if (c + 1 > maxCols) maxCols = c + 1;
                }
            }

            if (dateColIndex == -1) {
                throw new Exception("未找到指定的日期列: " + dateColumnName);
            }

            // 计算日期范围
            LocalDate today = LocalDate.now();
            LocalDate startDate;
            if (today.getDayOfMonth() <= 20) {
                // 前3个月：当天<=20号，往前推2个月到1号
                startDate = today.minusMonths(2).withDayOfMonth(1);
            } else {
                // 前2个月：当天>20号，往前推1个月到1号
                startDate = today.minusMonths(1).withDayOfMonth(1);
            }
            LocalDate endDate = today;

            // 读取所有数据行并分类
            List<Integer> matchedRows = new ArrayList<>();
            List<Integer> unmatchedRows = new ArrayList<>();
            DateTimeFormatter[] formatters = {
                DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                DateTimeFormatter.ofPattern("yyyy/MM/dd"),
                DateTimeFormatter.ofPattern("yyyy年MM月dd日"),
                DateTimeFormatter.ofPattern("yyyyMMdd"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
                DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss")
            };

            for (int r = 1; r <= lastRow; r++) {
                Row row = srcSheet.getRow(r);
                if (row == null) continue;

                Cell dateCell = row.getCell(dateColIndex);
                String dateStr = ExcelUtils.getCellStringValue(dateCell);
                if (dateStr == null || dateStr.isEmpty()) {
                    unmatchedRows.add(r);
                    continue;
                }

                LocalDate cellDate = parseDate(dateStr, formatters);
                if (cellDate == null) {
                    unmatchedRows.add(r);
                    continue;
                }

                if ((cellDate.isEqual(startDate) || cellDate.isAfter(startDate))
                        && (cellDate.isEqual(endDate) || cellDate.isBefore(endDate))) {
                    matchedRows.add(r);
                } else {
                    unmatchedRows.add(r);
                }
            }

            // 创建输出工作簿
            try (Workbook outWorkbook = new XSSFWorkbook()) {
                // 写入 "符合日期" 工作表
                writeDataSheet(outWorkbook, "符合日期", srcSheet, headerRow, matchedRows, maxCols);
                // 写入 "不符合日期" 工作表
                writeDataSheet(outWorkbook, "不符合日期", srcSheet, headerRow, unmatchedRows, maxCols);

                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outWorkbook.write(fos);
                }
            }
        }
    }

    private LocalDate parseDate(String dateStr, DateTimeFormatter[] formatters) {
        // 尝试先去掉时间部分（处理 "2026-05-15 00:00:00" 格式）
        String clean = dateStr.trim();
        if (clean.contains(" ")) {
            clean = clean.split(" ")[0];
        }
        for (DateTimeFormatter fmt : formatters) {
            try {
                return LocalDate.parse(clean, fmt);
            } catch (DateTimeParseException ignored) {
            }
        }
        // 尝试解析 Excel 序列号数字
        try {
            double serial = Double.parseDouble(clean);
            if (serial > 1) {
                return LocalDate.of(1899, 12, 30).plusDays((long) serial);
            }
        } catch (NumberFormatException ignored) {
        }
        return null;
    }

    private void writeDataSheet(Workbook workbook, String sheetName, Sheet srcSheet,
                                 Row headerRow, List<Integer> rowIndices, int colCount) {
        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle headerStyle = ExcelUtils.createHeaderStyle(workbook);

        // 复制表头
        Row newHeader = sheet.createRow(0);
        for (int c = 0; c < colCount; c++) {
            Cell srcCell = headerRow.getCell(c);
            Cell destCell = newHeader.createCell(c);
            if (srcCell != null) {
                ExcelUtils.copyCellValue(srcCell, destCell);
                destCell.setCellStyle(headerStyle);
            }
        }

        // 复制数据行（所有列为文本格式）
        int newRowIdx = 1;
        for (int rowIndex : rowIndices) {
            Row srcRow = srcSheet.getRow(rowIndex);
            if (srcRow == null) continue;
            Row destRow = sheet.createRow(newRowIdx++);
            for (int c = 0; c < colCount; c++) {
                Cell srcCell = srcRow.getCell(c);
                Cell destCell = destRow.createCell(c);
                if (srcCell != null) {
                    // 所有值转为字符串输出，避免科学计数法
                    String val = ExcelUtils.getCellStringValue(srcCell);
                    destCell.setCellValue(val);
                    destCell.setCellType(CellType.STRING);
                }
            }
        }
    }
}
