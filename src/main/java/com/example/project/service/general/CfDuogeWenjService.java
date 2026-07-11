package com.example.project.service.general;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 对应 Python: cf_duoge_wenj.py
 * 功能：按指定列将 Excel 拆分为多个独立的 Excel 文件（打包为 ZIP 返回）
 */
@Service
public class CfDuogeWenjService {

    /**
     * 按列拆分 Excel 为多个独立文件并打包为 ZIP
     * @param file       上传的 Excel 文件
     * @param colLetters 要拆分的列字母（多个用逗号分隔，如 "A,B"）
     * @return 生成的 ZIP 文件路径
     */
    public String splitToFiles(MultipartFile file, String colLetters, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }
        if (colLetters == null || colLetters.trim().isEmpty()) {
            throw new IllegalArgumentException("请指定要拆分的列字母");
        }

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            file.transferTo(inputFile.toPath());

            // 解析列字母
            String[] letters = colLetters.split(",");
            int[] colIndices = new int[letters.length];
            for (int i = 0; i < letters.length; i++) {
                colIndices[i] = ExcelUtils.excelColumnToIndex(letters[i].trim());
            }

            File zipFile = new File(outputDir, "split_files_" + System.currentTimeMillis() + ".zip");
            processSplitToFiles(inputFile, zipFile, colIndices);
            return zipFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) {
                inputFile.delete();
            }
        }
    }

    private void processSplitToFiles(File inputFile, File zipFile, int[] colIndices) throws Exception {
        // 读取源文件
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook srcWorkbook = new XSSFWorkbook(fis)) {

            Sheet srcSheet = srcWorkbook.getSheetAt(0);
            int lastRow = srcSheet.getLastRowNum();

            if (lastRow < 1) {
                throw new Exception("Excel 文件没有数据");
            }

            Row headerRow = srcSheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("Excel 文件没有表头行");
            }

            // 获取所有数据行
            List<Integer> allDataRows = new ArrayList<>();
            for (int r = 1; r <= lastRow; r++) {
                if (srcSheet.getRow(r) != null) {
                    allDataRows.add(r);
                }
            }

            if (allDataRows.isEmpty()) {
                throw new Exception("Excel 文件没有数据行");
            }

            // 用于存储所有生成的文件条目（文件名 → 行索引列表）
            Map<String, List<Integer>> fileEntries = new LinkedHashMap<>();

            // 对每个指定列的值进行组合，生成文件条目
            // 例如按 A 列和 B 列拆分，A=销售部、B=生产部 → 文件名为"销售部_生产部.xlsx"
            for (int rowIndex : allDataRows) {
                Row row = srcSheet.getRow(rowIndex);
                StringBuilder keyBuilder = new StringBuilder();
                for (int colIdx : colIndices) {
                    String val = ExcelUtils.getCellStringValue(row.getCell(colIdx));
                    if (val == null || val.isEmpty()) val = "空值";
                    if (keyBuilder.length() > 0) keyBuilder.append("_");
                    keyBuilder.append(val);
                }
                String key = keyBuilder.toString();
                // 清理文件名中的非法字符
                key = key.replaceAll("[\\\\/:*?\"<>|]", "_");
                if (key.length() > 100) {
                    key = key.substring(0, 100);
                }
                fileEntries.computeIfAbsent(key, k -> new ArrayList<>()).add(rowIndex);
            }

            // 计算最大列数
            int maxCols = 0;
            for (int r = 0; r <= lastRow; r++) {
                Row row = srcSheet.getRow(r);
                if (row != null && row.getLastCellNum() > maxCols) {
                    maxCols = row.getLastCellNum();
                }
            }
            int colCount = Math.max(maxCols, 1);

            // 生成 ZIP 文件
            try (FileOutputStream fos = new FileOutputStream(zipFile);
                 ZipOutputStream zos = new ZipOutputStream(fos)) {

                for (Map.Entry<String, List<Integer>> entry : fileEntries.entrySet()) {
                    String fileName = entry.getKey() + ".xlsx";
                    List<Integer> rowIndices = entry.getValue();

                    // 创建每个拆分文件的临时工作簿
                    try (Workbook splitWorkbook = new XSSFWorkbook()) {
                        Sheet splitSheet = splitWorkbook.createSheet("Data");

                        // 复制表头
                        Row newHeader = splitSheet.createRow(0);
                        ExcelUtils.copyRow(headerRow, newHeader, 0, colCount - 1);

                        // 复制数据行
                        int newRowIdx = 1;
                        for (int rowIndex : rowIndices) {
                            Row srcRow = srcSheet.getRow(rowIndex);
                            if (srcRow == null) continue;
                            Row destRow = splitSheet.createRow(newRowIdx++);
                            ExcelUtils.copyRow(srcRow, destRow, 0, colCount - 1);
                        }

                        // 将工作簿写入 ZIP
                        zos.putNextEntry(new ZipEntry(fileName));
                        splitWorkbook.write(zos);
                        zos.closeEntry();
                    }
                }
            }
        }
    }
}
