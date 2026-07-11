package com.example.project.service.hr;

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
 * 对应 Python: cf_zdnrfyq_nrcs_wsx.py
 * 功能：根据"拆分参数"工作表中的配置自动分组拆分 Excel
 *
 * 拆分参数工作表结构：
 *   A列 = 公司名称（数据表中拆分列的值）
 *   B列 = 分组编号（同一编号的公司分到同一个文件）
 *
 * 输出：{分组编号}_{原文件名}.xlsx
 */
@Service
public class CfZdnrfyqNrcsWsxService {

    private static final String PARAM_SHEET_NAME = "拆分参数";

    /**
     * 根据拆分参数自动分组拆分 Excel 文件
     * @param file        上传的 Excel 文件（必须包含"拆分参数"工作表）
     * @param colLetter   要拆分的列字母
     * @param outputDir   输出目录
     * @return ZIP 文件路径
     */
    public String splitByConfig(MultipartFile file, String colLetter, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            file.transferTo(inputFile.toPath());

            File zipFile = new File(outputDir, "config_split_" + System.currentTimeMillis() + ".zip");
            processConfigSplit(inputFile, zipFile, colLetter);
            return zipFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    private void processConfigSplit(File inputFile, File zipFile, String colLetter) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // 1. 检查并读取"拆分参数"工作表
            Sheet paramSheet = workbook.getSheet(PARAM_SHEET_NAME);
            if (paramSheet == null) {
                throw new Exception("Excel 文件中找不到'" + PARAM_SHEET_NAME + "'工作表");
            }

            // 2. 构建分组映射 {分组编号 → [公司名称列表]}
            Map<String, List<String>> groupMap = new LinkedHashMap<>();
            int paramLastRow = paramSheet.getLastRowNum();
            for (int r = 1; r <= paramLastRow; r++) {
                Row row = paramSheet.getRow(r);
                if (row == null) continue;

                String companyName = ExcelUtils.getCellStringValue(row.getCell(0)); // A列：公司名称
                String groupNum   = ExcelUtils.getCellStringValue(row.getCell(1)); // B列：分组编号

                if (companyName != null && !companyName.isEmpty()
                        && groupNum != null && !groupNum.isEmpty()) {
                    groupMap.computeIfAbsent(groupNum, k -> new ArrayList<>()).add(companyName);
                }
            }

            if (groupMap.isEmpty()) {
                throw new Exception("拆分参数工作表中没有有效的分组数据");
            }

            // 3. 获取数据工作表（第一个非"拆分参数"的工作表）
            Sheet dataSheet = null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (!workbook.getSheetName(i).equals(PARAM_SHEET_NAME)) {
                    dataSheet = workbook.getSheetAt(i);
                    break;
                }
            }
            if (dataSheet == null) {
                throw new Exception("找不到数据工作表");
            }

            // 4. 获取拆分列的索引
            int splitColIndex = ExcelUtils.excelColumnToIndex(colLetter);
            Row headerRow = dataSheet.getRow(0);
            if (headerRow == null) {
                throw new Exception("数据工作表没有表头行");
            }

            int lastDataRow = dataSheet.getLastRowNum();
            if (lastDataRow < 1) {
                throw new Exception("数据工作表没有数据");
            }

            // 计算最大列数
            int maxCols = 0;
            for (int r = 0; r <= lastDataRow; r++) {
                Row row = dataSheet.getRow(r);
                if (row != null && row.getLastCellNum() > maxCols) {
                    maxCols = row.getLastCellNum();
                }
            }
            int colCount = Math.max(maxCols, splitColIndex + 1);

            // 5. 按分组筛选数据行
            // 构建反向映射：公司名称 → 分组编号
            Map<String, String> companyToGroup = new HashMap<>();
            for (Map.Entry<String, List<String>> entry : groupMap.entrySet()) {
                for (String company : entry.getValue()) {
                    companyToGroup.put(company, entry.getKey());
                }
            }

            // 遍历数据行，按公司名分组
            Map<String, List<Integer>> resultGroups = new LinkedHashMap<>();
            for (int r = 1; r <= lastDataRow; r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;

                String cellVal = ExcelUtils.getCellStringValue(row.getCell(splitColIndex));
                if (cellVal == null || cellVal.isEmpty()) continue;

                String groupId = companyToGroup.get(cellVal);
                if (groupId != null) {
                    resultGroups.computeIfAbsent(groupId, k -> new ArrayList<>()).add(r);
                }
            }

            if (resultGroups.isEmpty()) {
                throw new Exception("没有匹配任何分组的数据");
            }

            // 6. 生成原文件名（不含扩展名）
            String originalName = inputFile.getName();
            if (originalName.contains(".")) {
                originalName = originalName.substring(0, originalName.lastIndexOf('.'));
            }
            // 去掉临时文件的前缀
            if (originalName.startsWith("input_")) {
                originalName = "data";
            }

            // 7. 写入 ZIP 文件
            try (FileOutputStream fos = new FileOutputStream(zipFile);
                 ZipOutputStream zos = new ZipOutputStream(fos)) {

                for (Map.Entry<String, List<Integer>> entry : resultGroups.entrySet()) {
                    String groupId = entry.getKey();
                    List<Integer> rowIndices = entry.getValue();

                    String fileName = groupId + "_" + originalName + ".xlsx";
                    fileName = fileName.replaceAll("[<>:\"/\\\\|?*]", "_");

                    try (Workbook splitWb = new XSSFWorkbook()) {
                        Sheet splitSheet = splitWb.createSheet("Data");

                        // 复制表头
                        Row newHeader = splitSheet.createRow(0);
                        ExcelUtils.copyRow(headerRow, newHeader, 0, colCount - 1);

                        // 复制数据行
                        int newIdx = 1;
                        for (int rowIndex : rowIndices) {
                            Row srcRow = dataSheet.getRow(rowIndex);
                            if (srcRow == null) continue;
                            Row destRow = splitSheet.createRow(newIdx++);
                            for (int c = 0; c < colCount; c++) {
                                Cell srcCell = srcRow.getCell(c);
                                Cell destCell = destRow.createCell(c);
                                if (srcCell != null) {
                                    String val = ExcelUtils.getCellStringValue(srcCell);
                                    destCell.setCellValue(val);
                                    // 长数字文本保护
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
}
