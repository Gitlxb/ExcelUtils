package com.example.project.service.finance;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * 对应 Python: cw_gys_qzy.py
 * 功能：供应商 Excel 文件合并
 * 处理逻辑：
 * 1. 从ZIP中读取所有Excel文件，跳过表头行（第1行）
 * 2. 添加"来源文件"列追踪来源
 * 3. 删除D列中的"供应商-"前缀
 * 4. 删除A列为空的行
 * 5. 处理"序号"列（保留第一行的"序号"，删除其余）
 */
@Service
public class CwGysQzyService {

    /**
     * 合并供应商 Excel 文件
     * @param file      上传的 ZIP 文件（内含多个 .xlsx 文件）
     * @param sheetName 要读取的工作表名称
     * @param outputDir 输出目录
     * @return 合并后的文件路径
     */
    public String mergeExcelFiles(MultipartFile file, String sheetName, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }
        if (sheetName == null || sheetName.trim().isEmpty()) {
            throw new IllegalArgumentException("请指定工作表名称");
        }

        File tempDir = new File(outputDir, "temp_merge_" + System.currentTimeMillis());
        File outputDirPath = new File(outputDir);
        if (!outputDirPath.exists()) outputDirPath.mkdirs();

        try {
            tempDir.mkdirs();

            // 解压 ZIP 到临时目录
            List<String> excelFiles = new ArrayList<>();
            try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(file.getBytes()))) {
                ZipEntry entry;
                while ((entry = zis.getNextEntry()) != null) {
                    if (!entry.isDirectory() && entry.getName().matches(".*\\.(xlsx|xls)$")) {
                        File destFile = new File(tempDir, new File(entry.getName()).getName());
                        Files.copy(zis, destFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                        excelFiles.add(destFile.getAbsolutePath());
                    }
                    zis.closeEntry();
                }
            }

            if (excelFiles.isEmpty()) {
                throw new Exception("ZIP文件中没有找到Excel文件");
            }

            // 执行合并
            File outputFile = new File(outputDir, "合并结果_" + System.currentTimeMillis() + ".xlsx");
            processMerge(excelFiles, sheetName, outputFile);
            return outputFile.getAbsolutePath();

        } finally {
            deleteDirectory(tempDir);
        }
    }

    private void processMerge(List<String> excelFiles, String sheetName, File outputFile) throws Exception {
        // 用于存储所有合并行的数据（每行是一个Cell值列表）
        List<List<String>> mergedRows = new ArrayList<>();

        for (String filePath : excelFiles) {
            String fileName = new File(filePath).getName();

            try (FileInputStream fis = new FileInputStream(filePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    // 尝试获取第一个工作表
                    if (workbook.getNumberOfSheets() > 0) {
                        sheet = workbook.getSheetAt(0);
                    }
                }
                if (sheet == null) continue;

                int lastRow = sheet.getLastRowNum();

                // 从第2行（索引1）开始读取，跳过表头
                for (int r = 1; r <= lastRow; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    int maxCol = Math.max(row.getLastCellNum(), 5); // 至少5列（A-D+来源文件）
                    List<String> rowData = new ArrayList<>();

                    // 读取 A 列到 D 列
                    for (int c = 0; c < 4; c++) {
                        Cell cell = row.getCell(c);
                        String val = ExcelUtils.getCellStringValue(cell);
                        rowData.add(val != null ? val : "");
                    }

                    // 检查 A 列是否为空（索引0）
                    String colA = rowData.get(0);
                    if (colA == null || colA.trim().isEmpty()) {
                        continue; // 删除A列为空的行
                    }

                    // 删除 D 列（索引3）中的"供应商-"前缀
                    String colD = rowData.get(3);
                    if (colD != null) {
                        colD = colD.replace("供应商-", "");
                        rowData.set(3, colD);
                    }

                    // 添加"来源文件"列
                    rowData.add(fileName);

                    mergedRows.add(rowData);
                }

            } catch (Exception e) {
                System.err.println("读取文件 " + fileName + " 时出错: " + e.getMessage());
            }
        }

        if (mergedRows.isEmpty()) {
            throw new Exception("没有找到有效数据");
        }

        // 处理"序号"列（A列，索引0）
        // 保留第一行的"序号"，删除其他行中包含"序号"的行
        if (!mergedRows.isEmpty()) {
            String firstColA = mergedRows.get(0).get(0);
            if (firstColA != null && firstColA.contains("序号")) {
                mergedRows.removeIf(row -> {
                    String aVal = row.get(0);
                    return aVal != null && aVal.contains("序号");
                });
                // 把第一行的"序号"加回去
                List<String> firstRow = new ArrayList<>();
                firstRow.add("序号");
                for (int i = 1; i < 4; i++) firstRow.add("");
                firstRow.add("来源文件");
                mergedRows.add(0, firstRow);
            }
        }

        // 保存结果
        try (Workbook outWorkbook = new XSSFWorkbook()) {
            Sheet outSheet = outWorkbook.createSheet("合并结果");

            for (int r = 0; r < mergedRows.size(); r++) {
                Row outRow = outSheet.createRow(r);
                List<String> rowData = mergedRows.get(r);
                for (int c = 0; c < rowData.size(); c++) {
                    Cell cell = outRow.createCell(c);
                    cell.setCellValue(rowData.get(c) != null ? rowData.get(c) : "");
                    cell.setCellType(CellType.STRING);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outWorkbook.write(fos);
            }
        }
    }

    private void deleteDirectory(File dir) {
        if (dir == null || !dir.exists()) return;
        File[] files = dir.listFiles();
        if (files != null) {
            for (File f : files) {
                if (f.isDirectory()) deleteDirectory(f);
                else f.delete();
            }
        }
        dir.delete();
    }
}
