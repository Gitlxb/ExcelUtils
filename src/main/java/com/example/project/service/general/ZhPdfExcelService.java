package com.example.project.service.general;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

/**
 * 对应 Python: zh_pdf_Excel.py
 * 功能：提取 PDF 中的表格和文本，转换为 Excel
 */
@Service
public class ZhPdfExcelService {

    /**
     * 将 PDF 文件转换为 Excel
     * @param file      上传的 PDF 文件
     * @param outputDir 输出目录
     */
    public String convertPdfToExcel(MultipartFile file, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("文件不能为空");
        }

        File inputFile = new File(outputDir, "input_" + System.currentTimeMillis() + ".pdf");
        file.transferTo(inputFile);

        try {
            File outputFile = new File(outputDir, "pdf_export_" + System.currentTimeMillis() + ".xlsx");

            try (PDDocument document = PDDocument.load(inputFile);
                 Workbook wb = new XSSFWorkbook();
                 FileOutputStream fos = new FileOutputStream(outputFile)) {

                PDFTextStripper stripper = new PDFTextStripper();
                stripper.setSortByPosition(true);

                int totalPages = document.getNumberOfPages();
                Sheet sheet = wb.createSheet("PDF内容");

                int rowIdx = 0;
                // 创建表头
                Row headerRow = sheet.createRow(rowIdx++);
                headerRow.createCell(0).setCellValue("内容");
                headerRow.createCell(1).setCellValue("数据类型");
                headerRow.createCell(2).setCellValue("来源页码");

                for (int pageNum = 1; pageNum <= totalPages; pageNum++) {
                    stripper.setStartPage(pageNum);
                    stripper.setEndPage(pageNum);
                    String text = stripper.getText(document);

                    if (text != null && !text.trim().isEmpty()) {
                        String[] lines = text.split("\\r?\\n");
                        for (String line : lines) {
                            line = line.trim();
                            if (line.isEmpty()) continue;

                            Row row = sheet.createRow(rowIdx++);
                            row.createCell(0).setCellValue(line);
                            row.createCell(1).setCellValue("文本");
                            row.createCell(2).setCellValue(pageNum);
                        }

                        // 添加分隔行
                        Row sepRow = sheet.createRow(rowIdx++);
                        sepRow.createCell(0).setCellValue("--- 第 " + pageNum + " 页结束 ---");
                        sepRow.createCell(1).setCellValue("分隔线");
                        sepRow.createCell(2).setCellValue(pageNum);
                    }
                }

                // 自动调整列宽
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);

                wb.write(fos);
            }

            return outputFile.getAbsolutePath();
        } finally {
            inputFile.delete();
        }
    }
}
