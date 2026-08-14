package com.example.project.service.finance.cwgjx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 财务工具箱 - 预支处理
 * 基于 yuzhi_frame.py 重构，含 8 条手续费规则
 */
@Service
public class CwgjxYuzhiService {

    public String process(MultipartFile file, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new Exception("文件不能为空");
        }
        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            file.transferTo(inputFile.toPath());
            File outputFile = new File(outputDir, "预支处理结果.xlsx");
            processData(inputFile, outputFile, null);
            return outputFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    private void processData(File inputFile, File outputFile, ProgressCallback callback) throws Exception {
        if (callback != null) callback.updateProgress(5, "正在读取 Excel 文件...");

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            if (lastRow == 0) throw new Exception("Excel文件没有数据");

            Map<String, Integer> colMap = new HashMap<>();
            for (int i = 0; i < 12; i++) colMap.put(String.valueOf((char) ('A' + i)), i);
            int colL = 11, colC = 2, colH = 7, colI = 8, colE = 4, colD = 3, colF = 5, colG = 6, colJ = 9, colK = 10;

            if (callback != null) callback.updateProgress(15, "正在应用条件规则...");

            // L列去空格
            for (int r = 1; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    Cell cell = row.getCell(colL);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String v = cell.getStringCellValue();
                        if (v != null) cell.setCellValue(v.replace(" ", ""));
                    }
                }
            }

            // 6条条件规则
            for (int r = 1; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String c = getStr(row, colC), h = getStr(row, colH), i = getStr(row, colI);
                boolean zero = false;
                if (h.contains("招商银行") && !c.contains("伟明")) zero = true;
                if (i.contains("驻厂")) zero = true;
                if (c.contains("伟明") && h.contains("中信银行")) zero = true;
                if (c.contains("伟明") && h.contains("招商银行") && i.contains("垫付")) zero = true;
                if (c.contains("瑞立") && (h.contains("招商银行") || h.contains("邮政"))) zero = true;
                if (c.contains("中世")) zero = true;
                // 规则7: 温州长江汽车电子 + 民生银行 -> 免手续费；其余银行 -> 5元
                if (c.contains("温州长江汽车电子") && h.contains("民生银行")) zero = true;
                if (c.contains("温州长江汽车电子") && !h.contains("民生银行")) {
                    Cell e = row.getCell(colE);
                    if (e == null) e = row.createCell(colE);
                    e.setCellValue(5);
                }
                // 规则8: 中铭工程 + 招商银行/垫付 -> 免手续费
                if (c.contains("中铭工程") && (h.contains("招商银行") || i.contains("垫付"))) zero = true;
                if (zero) {
                    Cell e = row.getCell(colE);
                    if (e == null) e = row.createCell(colE);
                    e.setCellValue(0);
                }
            }

            if (callback != null) callback.updateProgress(35, "正在计算 E 列数值...");
            for (int r = 1; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell e = row.getCell(colE);
                if (e == null || e.getCellType() == CellType.BLANK) {
                    double d = getDbl(row, colD);
                    String c = getStr(row, colC);
                    if (e == null) e = row.createCell(colE);
                    // 规则8: 中铭工程 特殊费率（仅对未命中免扣条件的行生效）
                    if (c.contains("中铭工程")) {
                        e.setCellValue(d <= 1000 ? 10 : d * 0.01);
                    } else {
                        e.setCellValue(d <= 500 ? 5 : d * 0.01);
                    }
                }
            }

            if (callback != null) callback.updateProgress(45, "正在计算 F 列数值...");
            for (int r = 1; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                double d = getDbl(row, colD), e = getDbl(row, colE);
                Cell f = row.getCell(colF);
                if (f == null) f = row.createCell(colF);
                f.setCellValue(d - e);
            }

            if (callback != null) callback.updateProgress(50, "正在保存文件...");
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
        }

        if (callback != null) callback.updateProgress(60, "正在拆分工作表...");

        // 重新打开按G列拆分
        try (FileInputStream fis = new FileInputStream(outputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            int colG = 6;

            Map<String, List<Integer>> map = new LinkedHashMap<>();
            for (int r = 1; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String g = getStr(row, colG);
                if (!g.isEmpty()) map.computeIfAbsent(g, k -> new ArrayList<>()).add(r);
            }
            if (map.isEmpty()) {
                List<Integer> all = new ArrayList<>();
                for (int r = 1; r <= lastRow; r++) if (sheet.getRow(r) != null) all.add(r);
                if (!all.isEmpty()) map.put("汇总", all);
                else throw new Exception("数据行为空，无法处理");
            }

            Workbook newWb = new XSSFWorkbook();
            CellStyle headerStyle = newWb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Row headerRow = sheet.getRow(0);
            AtomicInteger idx = new AtomicInteger(0);
            int total = map.size();

            for (Map.Entry<String, List<Integer>> e : map.entrySet()) {
                String name = e.getKey();
                if (name.length() > 31) name = name.substring(0, 31);
                Sheet ns = newWb.createSheet(name);
                Row nh = ns.createRow(0);
                if (headerRow != null) {
                    for (int c = 0; c <= 11; c++) {
                        Cell sc = headerRow.getCell(c);
                        Cell tc = nh.createCell(c);
                        if (sc != null) { copyCell(sc, tc); tc.setCellStyle(headerStyle); }
                    }
                }
                int nr = 1;
                for (int rowIdx : e.getValue()) {
                    Row sr = sheet.getRow(rowIdx);
                    if (sr == null) continue;
                    Row tr = ns.createRow(nr++);
                    for (int c = 0; c <= 11; c++) {
                        Cell sc = sr.getCell(c);
                        if (sc != null) copyCell(sc, tr.createCell(c));
                    }
                }
                int p = 60 + (idx.incrementAndGet() * 10 / Math.max(total, 1));
                if (callback != null) callback.updateProgress(p, "正在拆分工作表：" + name);
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                newWb.write(fos);
            }
            newWb.close();
        }

        if (callback != null) callback.updateProgress(70, "正在加载工作簿并格式化...");

        // 格式化
        try (FileInputStream fis = new FileInputStream(outputFile);
             Workbook wb = new XSSFWorkbook(fis)) {

            CellStyle yellow = wb.createCellStyle();
            yellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CellStyle dupRed = wb.createCellStyle();
            dupRed.setFillForegroundColor(IndexedColors.RED.getIndex());
            dupRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            int wsIdx = 0;
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                Sheet cs = wb.getSheetAt(s);
                int sr = cs.getLastRowNum();

                // 确定合计行位置（只查找，此时不创建）
                int totalRow = sr + 1;
                for (int r = 1; r <= sr; r++) {
                    Row row = cs.getRow(r);
                    if (row == null || row.getCell(2) == null || row.getCell(2).getCellType() == CellType.BLANK) {
                        totalRow = r;
                        break;
                    }
                }

                // 标黄
                for (int r = 1; r <= sr; r++) {
                    Row row = cs.getRow(r);
                    if (row == null) continue;
                    String iVal = getStr(row, 8), cVal = getStr(row, 2), hVal = getStr(row, 7);
                    if (iVal.contains("驻厂")) for (int c : new int[]{3,4,5,6,7,8}) setStyle(row, c, yellow);
                    if (!(cVal.contains("伟明") && hVal.contains("招商银行")) && hVal.contains("招商银行"))
                        for (int c : new int[]{3,4,5,6,7}) setStyle(row, c, yellow);
                    if (cVal.contains("伟明") && hVal.contains("中信银行"))
                        for (int c : new int[]{3,4,5,6,7}) setStyle(row, c, yellow);
                    if (cVal.contains("瑞立") && (hVal.contains("招商银行") || hVal.contains("邮政")))
                        for (int c : new int[]{3,4,5,6,7}) setStyle(row, c, yellow);
                    if (cVal.contains("中世"))
                        for (int c : new int[]{3,4,5,6,7}) setStyle(row, c, yellow);
                    // 规则7标黄: 温州长江汽车电子 + 民生银行 -> 黄色标记 D/E/F/G/H 列
                    if (cVal.contains("温州长江汽车电子") && hVal.contains("民生银行"))
                        for (int c : new int[]{3,4,5,6,7}) setStyle(row, c, yellow);
                    // 规则8标黄: 中铭工程 + 招商银行 -> D/E/F/G/H; 中铭工程 + 垫付 -> D/E/F/G/H/I
                    if (cVal.contains("中铭工程") && hVal.contains("招商银行"))
                        for (int c : new int[]{3,4,5,6,7}) setStyle(row, c, yellow);
                    if (cVal.contains("中铭工程") && iVal.contains("垫付"))
                        for (int c : new int[]{3,4,5,6,7,8}) setStyle(row, c, yellow);
                }

                // J/K重复值
                Map<String, Integer> jc = new HashMap<>(), kc = new HashMap<>();
                for (int r = 1; r <= sr; r++) {
                    Row row = cs.getRow(r); if (row == null) continue;
                    String j = getStr(row, 9), k = getStr(row, 10);
                    if (!j.isEmpty()) jc.put(j, jc.getOrDefault(j, 0) + 1);
                    if (!k.isEmpty()) kc.put(k, kc.getOrDefault(k, 0) + 1);
                }
                for (int r = 1; r <= sr; r++) {
                    Row row = cs.getRow(r); if (row == null) continue;
                    String j = getStr(row, 9), k = getStr(row, 10);
                    if (!j.isEmpty() && jc.getOrDefault(j, 0) > 1) setStyle(row, 9, dupRed);
                    if (!k.isEmpty() && kc.getOrDefault(k, 0) > 1) setStyle(row, 10, dupRed);
                }

                // 账号/户名/金额/汇款备注（先汇总，此时F列仍为数值）
                int start = totalRow + 5;
                Row ih = cs.createRow(start);
                ih.createCell(2).setCellValue("账号");
                ih.createCell(3).setCellValue("户名");
                ih.createCell(4).setCellValue("金额");
                ih.createCell(5).setCellValue("汇款备注");
                int dr = start + 1;
                double sum = 0;
                for (int r = 1; r < totalRow; r++) {
                    Row row = cs.getRow(r); if (row == null) continue;
                    String bc = getStr(row, 11), an = getStr(row, 10), name = getStr(row, 9);
                    double amt = getDbl(row, 5);
                    Row nr = cs.createRow(dr++);
                    nr.createCell(2).setCellValue(bc);
                    nr.createCell(3).setCellValue(an);
                    nr.createCell(4).setCellValue(amt);
                    nr.createCell(5).setCellValue(name);
                    sum += amt;
                }
                Row at = cs.createRow(dr);
                at.createCell(3).setCellValue("金额合计");
                Cell ac = at.createCell(4);
                ac.setCellValue(sum); ac.setCellStyle(yellow);

                // 给数据行F列添加公式 =D-E
                for (int r = 1; r < totalRow; r++) {
                    Row row = cs.getRow(r);
                    if (row == null) continue;
                    Cell f = row.getCell(5);
                    if (f == null) f = row.createCell(5);
                    f.setCellFormula("D" + (r + 1) + "-E" + (r + 1));
                }

                // 合计行
                Row totalRowObj = cs.getRow(totalRow);
                if (totalRowObj == null) totalRowObj = cs.createRow(totalRow);
                totalRowObj.createCell(2).setCellValue("合计");
                Cell dTotal = totalRowObj.createCell(3);
                dTotal.setCellFormula("SUM(D2:D" + totalRow + ")");
                Cell eTotal = totalRowObj.createCell(4);
                eTotal.setCellFormula("SUM(E2:E" + totalRow + ")");
                eTotal.setCellStyle(yellow);
                Cell fTotal = totalRowObj.createCell(5);
                fTotal.setCellFormula("SUM(F2:F" + totalRow + ")");
                fTotal.setCellStyle(yellow);

                int p = 75 + ((wsIdx + 1) * 20 / wb.getNumberOfSheets());
                if (callback != null) callback.updateProgress(p, "正在处理工作表：" + cs.getSheetName());
                wsIdx++;
            }

            if (callback != null) callback.updateProgress(95, "正在保存最终文件...");
            wb.setForceFormulaRecalculation(true);
            try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
        }
        if (callback != null) callback.updateProgress(100, "处理完成！");
    }

    private String getStr(Row row, int col) {
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }

    private double getDbl(Row row, int col) {
        Cell cell = row.getCell(col);
        if (cell == null) return 0;
        try {
            switch (cell.getCellType()) {
                case NUMERIC: return cell.getNumericCellValue();
                case STRING: String s = cell.getStringCellValue(); return s.isEmpty() ? 0 : Double.parseDouble(s);
                case FORMULA: return cell.getNumericCellValue();
                default: return 0;
            }
        } catch (Exception e) { return 0; }
    }

    private void copyCell(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING: dst.setCellValue(src.getStringCellValue()); break;
            case NUMERIC: dst.setCellValue(src.getNumericCellValue()); break;
            case BOOLEAN: dst.setCellValue(src.getBooleanCellValue()); break;
            case FORMULA: dst.setCellFormula(src.getCellFormula()); break;
            default: dst.setCellType(CellType.BLANK); break;
        }
    }

    private void setStyle(Row row, int col, CellStyle style) {
        Cell cell = row.getCell(col);
        if (cell == null) cell = row.createCell(col);
        cell.setCellStyle(style);
    }

    public interface ProgressCallback {
        void updateProgress(int percent, String message);
    }
}
