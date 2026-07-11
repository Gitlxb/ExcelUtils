package com.example.project.service.general;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

/**
 * 对应 Python: cf_biaoge_li.py
 * 功能：按指定列将 Excel 拆分为同一个文件中的多个工作表
 */
@Service
public class CfBiaogeLiService {

    /**
     * 按列拆分 Excel 为多个工作表
     * @param file      上传的 Excel 文件
     * @param colLetter 要拆分的列字母（如 A, B, AA）
     * @return 处理后的文件路径
     */
    public String splitToSheets(MultipartFile file, String colLetter, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            file.transferTo(inputFile.toPath());
            File outputFile = new File(outputDir, "split_sheets_" + System.currentTimeMillis() + ".xlsx");
            processSplit(inputFile, outputFile, colLetter);
            return outputFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) {
                inputFile.delete();
            }
        }
    }

    private void processSplit(File inputFile, File outputFile, String colLetter) throws Exception {
        // 读取源文件
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook srcWorkbook = new XSSFWorkbook(fis)) {

            Sheet srcSheet = srcWorkbook.getSheetAt(0);
            int lastRow = srcSheet.getLastRowNum();

            if (lastRow < 1) {
                throw new Exception("Excel 文件没有数据");
            }

            // 获取列索引
            int splitColIndex = ExcelUtils.excelColumnToIndex(colLetter);

            // 读取表头
            Row headerRow = srcSheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("Excel 文件没有表头行");
            }

            // 计算最大列数
            int maxCols = 0;
            for (int r = 0; r <= lastRow; r++) {
                Row row = srcSheet.getRow(r);
                if (row != null && row.getLastCellNum() > maxCols) {
                    maxCols = row.getLastCellNum();
                }
            }
            int colCount = Math.max(maxCols, splitColIndex + 1);

            // 按指定列分组
            Map<String, List<Integer>> groupMap = new LinkedHashMap<>();
            for (int rowIndex = 1; rowIndex <= lastRow; rowIndex++) {
                Row row = srcSheet.getRow(rowIndex);
                if (row == null) continue;

                Cell cell = row.getCell(splitColIndex);
                String cellValue = ExcelUtils.getCellStringValue(cell);

                if (cellValue == null || cellValue.isEmpty()) {
                    cellValue = "空值";
                }

                groupMap.computeIfAbsent(cellValue, k -> new ArrayList<>()).add(rowIndex);
            }

            // 创建输出工作簿
            try (Workbook outWorkbook = new XSSFWorkbook()) {
                CellStyle headerStyle = ExcelUtils.createHeaderStyle(outWorkbook);

                // 1. 创建"原始表" - 包含所有原始数据
                Sheet originalSheet = outWorkbook.createSheet("原始表");
                copyHeaderWithStyle(headerRow, originalSheet, 0, colCount, headerStyle);
                int origRowIdx = 1;
                for (int rowIndex = 1; rowIndex <= lastRow; rowIndex++) {
                    Row srcRow = srcSheet.getRow(rowIndex);
                    if (srcRow == null) continue;
                    Row destRow = originalSheet.createRow(origRowIdx++);
                    ExcelUtils.copyRow(srcRow, destRow, 0, colCount - 1);
                }

                // 2. 为每个分组创建独立工作表
                Set<String> usedSheetNames = new HashSet<>();
                for (Map.Entry<String, List<Integer>> entry : groupMap.entrySet()) {
                    String groupValue = entry.getKey();
                    List<Integer> rowIndices = entry.getValue();

                    // 生成唯一的工作表名称
                    String sheetName = ExcelUtils.getUniqueSheetName(groupValue, usedSheetNames);
                    usedSheetNames.add(sheetName);

                    Sheet newSheet = outWorkbook.createSheet(sheetName);

                    // 复制表头
                    copyHeaderWithStyle(headerRow, newSheet, 0, colCount, headerStyle);

                    // 复制数据行
                    int newRowIdx = 1;
                    for (int rowIndex : rowIndices) {
                        Row srcRow = srcSheet.getRow(rowIndex);
                        if (srcRow == null) continue;
                        Row destRow = newSheet.createRow(newRowIdx++);
                        ExcelUtils.copyRow(srcRow, destRow, 0, colCount - 1);
                    }
                }

                // 保存
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    outWorkbook.write(fos);
                }
            }
        }
    }

    private void copyHeaderWithStyle(Row headerRow, Sheet sheet, int rowIndex, int colCount, CellStyle style) {
        Row destRow = sheet.createRow(rowIndex);
        for (int c = 0; c < colCount; c++) {
            Cell srcCell = headerRow.getCell(c);
            Cell destCell = destRow.createCell(c);
            if (srcCell != null) {
                ExcelUtils.copyCellValue(srcCell, destCell);
                destCell.setCellStyle(style);
            }
        }
    }
}
