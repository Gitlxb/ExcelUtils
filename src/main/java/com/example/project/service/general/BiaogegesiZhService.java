package com.example.project.service.general;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * 对应 Python: biaogegesi_zh.py
 * 功能：将 .xls/.csv/.xlsm/.xlsb 等格式批量转换为 .xlsx
 */
@Service
public class BiaogegesiZhService {

    /**
     * 将上传的 ZIP 包内所有支持的文件转换为 .xlsx，再打包返回
     */
    public String convertToXlsx(MultipartFile zipFile, String outputDir) throws Exception {
        if (zipFile == null || zipFile.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }

        File tempDir = new File(outputDir, "temp_convert_" + System.currentTimeMillis());
        File outDir = new File(outputDir);
        if (!outDir.exists()) outDir.mkdirs();

        try {
            tempDir.mkdirs();

            // 解压 ZIP
            try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(zipFile.getBytes()))) {
                ZipEntry entry;
                while ((entry = zis.getNextEntry()) != null) {
                    if (!entry.isDirectory()) {
                        File dest = new File(tempDir, entry.getName());
                        dest.getParentFile().mkdirs();
                        Files.copy(zis, dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    }
                    zis.closeEntry();
                }
            }

            // 转换并打包
            File outputZip = new File(outputDir, "converted_xlsx_" + System.currentTimeMillis() + ".zip");
            try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(outputZip))) {
                convertDirectory(tempDir, tempDir, zos);
            }

            return outputZip.getAbsolutePath();
        } finally {
            deleteDirectory(tempDir);
        }
    }

    private void convertDirectory(File baseDir, File currentDir, ZipOutputStream zos) throws IOException {
        File[] files = currentDir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                convertDirectory(baseDir, file, zos);
                continue;
            }
            String name = file.getName();
            String lower = name.toLowerCase();
            String relativePath = baseDir.toPath().relativize(file.toPath()).toString();
            String zipEntryName = relativePath.substring(0, relativePath.lastIndexOf('.')) + ".xlsx";

            // 处理重名
            zipEntryName = resolveUniqueName(zos, zipEntryName);

            if (lower.endsWith(".xlsx") || lower.endsWith(".xlsm") || lower.endsWith(".xlsb")) {
                // 直接复制为 xlsx
                zos.putNextEntry(new ZipEntry(zipEntryName));
                Files.copy(file.toPath(), zos);
                zos.closeEntry();
            } else if (lower.endsWith(".xls")) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook src = new HSSFWorkbook(fis);
                     ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                    Workbook dst = new XSSFWorkbook();
                    copyWorkbook(src, dst);
                    dst.write(bos);
                    zos.putNextEntry(new ZipEntry(zipEntryName));
                    zos.write(bos.toByteArray());
                    zos.closeEntry();
                    dst.close();
                }
            } else if (lower.endsWith(".csv")) {
                try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(file), detectEncoding(file)));
                     ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                    Workbook dst = new XSSFWorkbook();
                    Sheet sheet = dst.createSheet("Sheet1");
                    String line;
                    int rowIdx = 0;
                    while ((line = br.readLine()) != null) {
                        Row row = sheet.createRow(rowIdx++);
                        String[] cols = line.split(",");
                        for (int i = 0; i < cols.length; i++) {
                            row.createCell(i).setCellValue(cols[i]);
                        }
                    }
                    dst.write(bos);
                    zos.putNextEntry(new ZipEntry(zipEntryName));
                    zos.write(bos.toByteArray());
                    zos.closeEntry();
                    dst.close();
                }
            }
        }
    }

    private String resolveUniqueName(ZipOutputStream zos, String name) {
        // ZipOutputStream 是流式的，无法在写入前检测重名。
        // 简单处理：如果目录结构不同但文件名相同，保留相对路径即可避免大部分冲突。
        return name;
    }

    private String detectEncoding(File file) throws IOException {
        // 简单检测 BOM
        try (FileInputStream fis = new FileInputStream(file)) {
            byte[] bom = new byte[3];
            int read = fis.read(bom);
            if (read >= 3 && (bom[0] & 0xFF) == 0xEF && (bom[1] & 0xFF) == 0xBB && (bom[2] & 0xFF) == 0xBF) {
                return "UTF-8";
            }
        }
        return "GBK";
    }

    private void copyWorkbook(Workbook src, Workbook dst) {
        for (int i = 0; i < src.getNumberOfSheets(); i++) {
            Sheet srcSheet = src.getSheetAt(i);
            Sheet dstSheet = dst.createSheet(srcSheet.getSheetName());
            copySheet(srcSheet, dstSheet);
        }
    }

    private void copySheet(Sheet src, Sheet dst) {
        for (Row srcRow : src) {
            Row dstRow = dst.createRow(srcRow.getRowNum());
            for (Cell srcCell : srcRow) {
                Cell dstCell = dstRow.createCell(srcCell.getColumnIndex());
                copyCell(srcCell, dstCell);
            }
        }
    }

    private void copyCell(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING:
                dst.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dst.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dst.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                dst.setCellFormula(src.getCellFormula());
                break;
            default:
                dst.setBlank();
                break;
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
