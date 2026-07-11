package com.example.project.service.finance;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.UUID;

/**
 * 批量批注工具（pl_cf_pizhu.py）
 * 功能1：批量添加批注（B~K列 → A列批注）
 * 功能2：批量拆分批注（A列批注 → 右侧列）
 */
@Service
public class PlCfPizhuService {

    /**
     * 功能1：批量添加批注
     * 遍历A列所有有值的单元格，将B~K列的非空内容用换行符拼接作为批注写入A列
     */
    public String addComments(MultipartFile file, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        try (Workbook wb = WorkbookFactory.create(new FileInputStream(tempFile))) {
            Sheet sheet = wb.getSheetAt(0);
            int count = 0;

            for (Row row : sheet) {
                Cell aCell = row.getCell(0);
                if (aCell == null) continue;

                String aValue = getCellString(aCell);
                if (aValue == null || aValue.isEmpty()) continue;

                // 收集B~K列（索引1~10）的非空内容
                StringBuilder sb = new StringBuilder();
                for (int col = 1; col <= 10; col++) {
                    Cell cell = row.getCell(col);
                    if (cell != null) {
                        String v = getCellString(cell);
                        if (v != null && !v.isEmpty()) {
                            if (sb.length() > 0) sb.append("\n");
                            sb.append(v);
                        }
                    }
                }

                String commentText = sb.toString();
                if (!commentText.isEmpty()) {
                    // 删除旧批注
                    if (aCell.getCellComment() != null) {
                        aCell.removeCellComment();
                    }
                    // 创建新批注
                    ClientAnchor anchor = wb.getCreationHelper().createClientAnchor();
                    anchor.setCol1(0);
                    anchor.setCol2(3);
                    anchor.setRow1(row.getRowNum());
                    anchor.setRow2(row.getRowNum() + 3);
                    Drawing<?> drawing = sheet.createDrawingPatriarch();
                    Comment comment = drawing.createCellComment(anchor);
                    comment.setString(wb.getCreationHelper().createRichTextString(commentText));
                    comment.setAuthor("BatchTool");
                    aCell.setCellComment(comment);
                    count++;
                }
            }

            String outputPath = uploadDir + File.separator + "批注添加结果.xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                wb.write(fos);
            }
            tempFile.delete();
            return outputPath;
        }
    }

    /**
     * 功能2：批量拆分批注
     * 遍历A列所有有批注的单元格，将批注内容按空格拆分，依次写入右侧列
     */
    public String splitComments(MultipartFile file, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        try (Workbook wb = WorkbookFactory.create(new FileInputStream(tempFile))) {
            Sheet sheet = wb.getSheetAt(0);
            int count = 0;

            for (Row row : sheet) {
                Cell aCell = row.getCell(0);
                if (aCell == null) continue;

                Comment comment = aCell.getCellComment();
                if (comment == null) continue;

                String commentText = comment.getString().getString();
                if (commentText == null || commentText.isEmpty()) continue;

                // 过滤掉以 "Admin:" 开头的行
                String[] lines = commentText.split("\n");
                StringBuilder filtered = new StringBuilder();
                for (String line : lines) {
                    if (line.trim().startsWith("Admin:")) continue;
                    if (filtered.length() > 0) filtered.append(" ");
                    filtered.append(line.trim());
                }

                // 规范化空格后按空格拆分
                String normalized = filtered.toString().replaceAll("\\s+", " ").trim();
                if (normalized.isEmpty()) continue;

                String[] splitValues = normalized.split(" ");
                for (int i = 0; i < splitValues.length; i++) {
                    Cell targetCell = row.getCell(1 + i);
                    if (targetCell == null) targetCell = row.createCell(1 + i);
                    targetCell.setCellValue(splitValues[i]);
                }
                count++;
            }

            String outputPath = uploadDir + File.separator + "批注拆分结果.xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                wb.write(fos);
            }
            tempFile.delete();
            return outputPath;
        }
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double d = cell.getNumericCellValue();
                if (d == (long) d) return String.valueOf((long) d);
                return String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default: return "";
        }
    }
}
