package com.example.project.service.hr;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 对应 Python: cf_gzbcf_hyp_simple.py（简化版）
 * 轻量快速的基础拆分功能，仅含 个税公式填充，无社保/公积金回填，无总计行。
 */
@Service
public class CfGzbcfHypSimpleService {

    private static final Set<String> FOOTER_KEYWORDS = Set.of("合计", "总计", "汇总", "小计", "平均");
    private static final List<String> SHEETS_TO_REMOVE = List.of("各公司支出", "未发工资的社保扣除人员");

    // ========================================================================
    //  公开 API（与增强版相同）
    // ========================================================================

    public List<String> getSheetNames(MultipartFile file) throws Exception {
        List<String> sheets = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) sheets.add(wb.getSheetName(i));
        }
        return sheets;
    }

    public List<String> getHeaders(MultipartFile file, String sheetName, int headerRow) throws Exception {
        List<String> headers = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) { if (wb.getNumberOfSheets() > 0) sheet = wb.getSheetAt(0); else return headers; }
            Row row = sheet.getRow(headerRow - 1);
            if (row == null) return headers;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                String val = getCellStringValue(row.getCell(c)).trim();
                if (!val.isEmpty()) headers.add(val);
            }
        }
        return headers;
    }

    public List<String> getColumnValues(MultipartFile file, String sheetName, String colName) throws Exception {
        List<String> values = new ArrayList<>();
        Set<String> seen = new LinkedHashSet<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = (sheetName != null && !sheetName.isEmpty()) ? wb.getSheet(sheetName) : null;
            if (sheet == null) sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(1);
            if (headerRow == null) return values;
            int colIdx = findColumnIndex(headerRow, colName);
            if (colIdx < 0) return values;
            int footerStart = detectFooterStart(sheet, colIdx);
            Map<String, String> mergeValues = buildMergeValueMap(sheet);
            String lastVal = "";
            for (int r = 2; r < footerStart; r++) {
                Row row = sheet.getRow(r);
                String val = "";
                if (row != null) {
                    String mergeKey = r + "_" + colIdx;
                    val = mergeValues.getOrDefault(mergeKey, getCellStringValue(row.getCell(colIdx)).trim());
                }
                if (val.isEmpty() && !lastVal.isEmpty()) val = lastVal;
                if (!val.isEmpty() && FOOTER_KEYWORDS.stream().noneMatch(val::contains)) {
                    if (seen.add(val)) values.add(val);
                }
                if (!val.isEmpty()) lastVal = val;
            }
        }
        return values;
    }

    // ========================================================================
    //  核心拆分
    // ========================================================================

    public String splitSalaryWithGroups(MultipartFile file, String originalName, String sheetName,
                                         String splitColName, List<List<String>> groups, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (splitColName == null || splitColName.trim().isEmpty()) throw new IllegalArgumentException("拆分依据列名不能为空");
        if (groups == null || groups.isEmpty()) return splitSalary(file, originalName, sheetName, splitColName, outputDir);

        File inputFile = new File(outputDir, "input_simple_" + System.currentTimeMillis() + ".xlsx");
        file.transferTo(inputFile);

        try (Workbook wb = WorkbookFactory.create(inputFile)) {
            Sheet sheet = resolveSheet(wb, sheetName);
            String originalSheetName = sheet.getSheetName();
            int splitColIdx = resolveSplitCol(sheet, splitColName);
            int footerStart = detectFooterStart(sheet, splitColIdx);
            int maxRow = sheet.getLastRowNum();
            if (footerStart > maxRow) footerStart = maxRow + 1;
            int maxCol = computeMaxCol(sheet, maxRow);
            Map<String, String> mergeValues = buildMergeValueMap(sheet);

            // 构建值到组的映射
            Map<String, Integer> valueToGroup = new LinkedHashMap<>();
            Map<Integer, String> groupNames = new LinkedHashMap<>();
            for (int gi = 0; gi < groups.size(); gi++) {
                List<String> gv = groups.get(gi);
                if (gv.isEmpty()) continue;
                groupNames.put(gi, sanitizeFilename(String.join("_", gv)));
                for (String v : gv) valueToGroup.put(v.trim(), gi);
            }

            Map<Integer, List<Integer>> groupRows = new LinkedHashMap<>();
            for (int i = 0; i < groups.size(); i++) groupRows.put(i, new ArrayList<>());
            Map<String, List<Integer>> ungrouped = new LinkedHashMap<>();
            String lastVal = "";

            for (int r = 2; r < footerStart; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String val = getSplitValue(row, r, splitColIdx, mergeValues);
                if (val.isEmpty() && !lastVal.isEmpty() && !isFooterRow(row)) val = lastVal;
                if (!val.isEmpty() && !isFooterRow(row)) {
                    Integer gi = valueToGroup.get(val);
                    if (gi != null) groupRows.get(gi).add(r);
                    else ungrouped.computeIfAbsent(val, k -> new ArrayList<>()).add(r);
                    lastVal = val;
                }
            }

            String zipName = extractBaseName(originalName) + "_拆分结果.zip";
            File zipFile = new File(outputDir, zipName);
            try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
                for (Map.Entry<Integer, List<Integer>> e : groupRows.entrySet()) {
                    if (e.getValue().isEmpty()) continue;
                    generateAndZip(inputFile, originalSheetName, groupNames.get(e.getKey()),
                            e.getValue(), footerStart, maxRow, maxCol, splitColIdx, zos, outputDir);
                }
                for (Map.Entry<String, List<Integer>> e : ungrouped.entrySet()) {
                    generateAndZip(inputFile, originalSheetName, e.getKey(),
                            e.getValue(), footerStart, maxRow, maxCol, splitColIdx, zos, outputDir);
                }
            }
            inputFile.delete();
            return zipFile.getAbsolutePath();
        }
    }

    public String splitSalary(MultipartFile file, String originalName, String sheetName,
                               String splitColName, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (splitColName == null || splitColName.trim().isEmpty()) throw new IllegalArgumentException("拆分依据列名不能为空");

        File inputFile = new File(outputDir, "input_simple_" + System.currentTimeMillis() + ".xlsx");
        file.transferTo(inputFile);

        try (Workbook wb = WorkbookFactory.create(inputFile)) {
            Sheet sheet = resolveSheet(wb, sheetName);
            String originalSheetName = sheet.getSheetName();
            int splitColIdx = resolveSplitCol(sheet, splitColName);
            int footerStart = detectFooterStart(sheet, splitColIdx);
            int maxRow = sheet.getLastRowNum();
            if (footerStart > maxRow) footerStart = maxRow + 1;
            int maxCol = computeMaxCol(sheet, maxRow);
            Map<String, String> mergeValues = buildMergeValueMap(sheet);

            Map<String, List<Integer>> groups = new LinkedHashMap<>();
            String lastVal = "";
            for (int r = 2; r < footerStart; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String val = getSplitValue(row, r, splitColIdx, mergeValues);
                // 连续的数据行允许继承上一个非空值（处理拆分列中合并或空单元格的情况）
                if (val.isEmpty() && !lastVal.isEmpty() && !isFooterRow(row)) val = lastVal;
                if (!val.isEmpty() && !isFooterRow(row)) {
                    groups.computeIfAbsent(val, k -> new ArrayList<>()).add(r);
                    lastVal = val;
                }
            }
            if (groups.isEmpty()) throw new IllegalArgumentException("未找到有效的拆分数据");

            System.out.println("===== [DIAG] 简化版 分组建行结果 =====");
            System.out.println("[DIAG] footerStart=" + footerStart + ", maxRow=" + maxRow + ", splitColIdx=" + splitColIdx);
            for (Map.Entry<String, List<Integer>> e : groups.entrySet()) {
                System.out.println("[DIAG] 组=[" + e.getKey() + "], 行数=" + e.getValue().size() + ", 行号=[" + e.getValue().get(0) + ".." + e.getValue().get(e.getValue().size()-1) + "]");
            }

            String zipName = extractBaseName(originalName) + "_拆分结果.zip";
            File zipFile = new File(outputDir, zipName);
            try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
                for (Map.Entry<String, List<Integer>> e : groups.entrySet()) {
                    generateAndZip(inputFile, originalSheetName, e.getKey(),
                            e.getValue(), footerStart, maxRow, maxCol, splitColIdx, zos, outputDir);
                }
            }
            return zipFile.getAbsolutePath();
        } finally {
            inputFile.delete();
        }
    }

    private void generateAndZip(File inputFile, String originalSheetName, String groupKey,
                                 List<Integer> rows, int footerStart, int maxRow, int maxCol,
                                 int splitColIdx, ZipOutputStream zos, String outputDir) throws Exception {
        String safeName = sanitizeFilename(groupKey);
        File deptFile = new File(outputDir, safeName + "_" + System.currentTimeMillis() + ".xlsx");
        Files.copy(inputFile.toPath(), deptFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        try {
            processDepartmentFile(deptFile, originalSheetName, groupKey, rows,
                    footerStart, maxRow, maxCol, splitColIdx);
            zos.putNextEntry(new ZipEntry(safeName + ".xlsx"));
            Files.copy(deptFile.toPath(), zos);
            zos.closeEntry();
        } finally {
            deptFile.delete();
        }
    }

    // ========================================================================
    //  核心：生成单个部门文件（简化版：无公积金/社保/总计行）
    // ========================================================================
    private void processDepartmentFile(File file, String originalSheetName, String deptName,
                                        List<Integer> dataRows, int footerStart, int maxRow,
                                        int maxCol, int splitColIdx) throws Exception {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet oldSheet = wb.getSheet(originalSheetName);
            if (oldSheet == null) oldSheet = wb.getSheetAt(0);

            // 1. 过滤幽灵行
            List<Integer> filteredRows = new ArrayList<>();
            for (int oldR : dataRows) {
                Row row = oldSheet.getRow(oldR);
                if (row == null) continue;
                boolean hasContent = false;
                for (int c = 0; c < maxCol; c++) {
                    if (c == splitColIdx) continue;
                    Cell cell = row.getCell(c);
                    if (cell != null && !getCellStringValue(cell).trim().isEmpty()) { hasContent = true; break; }
                }
                if (hasContent) filteredRows.add(oldR);
            }
            if (filteredRows.isEmpty()) throw new IllegalStateException("过滤后无有效数据行: " + deptName);

            // 2. 保留行列表：第1行(标题) + 第2行(表头) + 数据行 + 表尾
            List<Integer> keepRows = new ArrayList<>();
            keepRows.add(0);
            keepRows.add(1);
            keepRows.addAll(filteredRows);

            // 预先计算有内容的表尾行
            List<Integer> actualFooterRows = new ArrayList<>();
            for (int r = footerStart; r <= maxRow; r++) {
                Row row = oldSheet.getRow(r);
                if (row == null) continue;
                for (int c = 0; c < maxCol; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && !getCellStringValue(cell).trim().isEmpty()) {
                        actualFooterRows.add(r);
                        break;
                    }
                }
            }
            keepRows.addAll(actualFooterRows);

            System.out.println("[DIAG] processDepartmentFile for=[" + deptName + "]");
            System.out.println("[DIAG]   dataRows.size=" + dataRows.size() + ", filteredRows.size=" + filteredRows.size());
            System.out.println("[DIAG]   footerStart=" + footerStart + ", actualFooterRows=[" + (actualFooterRows.isEmpty() ? "empty" : actualFooterRows.get(0) + ".." + actualFooterRows.get(actualFooterRows.size()-1)) + "]");
            System.out.println("[DIAG]   keepRows=[" + keepRows.get(0) + "," + keepRows.get(1) + ",.." + keepRows.get(keepRows.size()-1) + "] total=" + keepRows.size());

            Set<Integer> keepSet = new HashSet<>(keepRows);
            Map<Integer, Integer> rowMap = new LinkedHashMap<>();
            int newR = 0;
            for (int oldR : keepRows) rowMap.put(oldR, newR++);
            int newDataLastRow = rowMap.get(filteredRows.get(filteredRows.size() - 1));
            int oldLastDataRow = footerStart - 1;

            // 3. 创建新 Sheet（先保留旧 Sheet 用于复制数据）
            String safeSheetName = getSafeSheetName(wb, sanitizeFilename(deptName));
            Sheet newSheet = wb.createSheet(safeSheetName);
            wb.setSheetOrder(safeSheetName, 0);

            // 4. 从旧 Sheet 逐行复制到新 Sheet（同一 Workbook，避免样式跨文件错误）
            for (int oldR : keepRows) {
                Row oldRow = oldSheet.getRow(oldR);
                Row newRow = newSheet.createRow(rowMap.get(oldR));
                if (oldRow == null) continue;
                for (int c = 0; c < maxCol; c++) {
                    Cell oldCell = oldRow.getCell(c);
                    if (oldCell == null) continue;
                    Cell newCell = newRow.createCell(c);
                    if (oldCell.getCellType() == CellType.FORMULA) {
                        newCell.setCellFormula(oldCell.getCellFormula());
                    } else {
                        copyCellValue(oldCell, newCell);
                    }
                    newCell.setCellStyle(oldCell.getCellStyle());
                }
                newRow.setHeight(oldRow.getHeight());
            }

            // 复制列宽
            for (int c = 0; c < maxCol; c++) newSheet.setColumnWidth(c, oldSheet.getColumnWidth(c));

            // 重建合并单元格
            for (int i = 0; i < oldSheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = oldSheet.getMergedRegion(i);
                Set<Integer> regionRows = new HashSet<>();
                for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) regionRows.add(r);
                if (keepSet.containsAll(regionRows)) {
                    newSheet.addMergedRegion(new CellRangeAddress(
                            rowMap.get(region.getFirstRow()), rowMap.get(region.getLastRow()),
                            region.getFirstColumn(), region.getLastColumn()));
                }
            }

            // 5. 删除旧 Sheet
            wb.removeSheetAt(wb.getSheetIndex(oldSheet));

            // 6. 调整所有公式（简化版占位符方式）
            int newMaxRow = newSheet.getLastRowNum();
            for (int r = 0; r <= newMaxRow; r++) {
                Row row = newSheet.getRow(r);
                if (row == null) continue;
                for (int c = 0; c < maxCol; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && cell.getCellType() == CellType.FORMULA) {
                        String f = cell.getCellFormula();
                        String adjusted = adjustFormulaSimple(f, rowMap, oldLastDataRow, newDataLastRow);
                        if (!f.equals(adjusted)) cell.setCellFormula(adjusted);
                    }
                }
            }

            // 7. 填充个税公式
            Map<String, List<Integer>> colNameMap = new LinkedHashMap<>();
            Row newHeader = newSheet.getRow(1);
            if (newHeader != null) {
                for (int c = 0; c < maxCol; c++) {
                    String name = getCellStringValue(newHeader.getCell(c)).trim();
                    if (!name.isEmpty()) colNameMap.computeIfAbsent(name, k -> new ArrayList<>()).add(c);
                }
            }
            List<Integer> taxCols = colNameMap.get("个税");
            if (taxCols != null && !taxCols.isEmpty()) {
                int taxCol = taxCols.get(0);
                for (int r = 2; r <= newDataLastRow; r++) {
                    Row row = newSheet.getRow(r);
                    Cell cell = row.getCell(taxCol);
                    if (cell == null) cell = row.createCell(taxCol);
                    int er = r + 1;
                    cell.setCellFormula(
                        "-ROUND(MAX((AK" + er + "-5000-AJ" + er + "-AT" + er
                        + "+AM" + er + "+AN" + er
                        + ")*{3,10,20,25,30,35,45}%-5*{0,42,282,532,882,1432,3032},),2)");
                }
            }

            // 8. 清理多余 Sheet
            for (String removeName : SHEETS_TO_REMOVE) {
                int idx = wb.getSheetIndex(removeName);
                if (idx >= 0) wb.removeSheetAt(idx);
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
    }

    // ========================================================================
    //  简化版公式调整（占位符保护跨Sheet引用）
    // ========================================================================
    private String adjustFormulaSimple(String formula, Map<Integer, Integer> rowMap,
                                        int oldLastDataRow, int newDataLastRow) {
        if (formula == null || formula.isEmpty()) return formula;

        // 占位符保护跨Sheet引用
        Pattern crossPattern = Pattern.compile(
                "('[^']+'|[A-Za-z0-9_\\u4e00-\\u9fa5]+)!" +
                "(\\$?[A-Za-z]{1,3}\\$?\\d+(:\\$?[A-Za-z]{1,3}\\$?\\d+)?)");
        List<String> crossPh = new ArrayList<>();
        Matcher crossMatcher = crossPattern.matcher(formula);
        while (crossMatcher.find()) crossPh.add(crossMatcher.group(0));

        String temp = formula;
        for (int i = crossPh.size() - 1; i >= 0; i--)
            temp = temp.replace(crossPh.get(i), "__CROSS" + i + "__");

        // 调整本Sheet行号
        Pattern localPattern = Pattern.compile("(\\$?[A-Za-z]{1,3}\\$?)(\\d+)");
        Matcher m = localPattern.matcher(temp);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String colPart = m.group(1);
            int rowNum = Integer.parseInt(m.group(2));
            int zeroBased = rowNum - 1;
            int newNum;
            if (rowMap.containsKey(zeroBased)) newNum = rowMap.get(zeroBased) + 1;
            else if (rowNum == oldLastDataRow + 1) newNum = newDataLastRow + 1;
            else newNum = rowNum;
            m.appendReplacement(sb, Matcher.quoteReplacement(colPart + newNum));
        }
        m.appendTail(sb);
        temp = sb.toString();

        // 还原跨Sheet引用
        for (int i = 0; i < crossPh.size(); i++)
            temp = temp.replace("__CROSS" + i + "__", crossPh.get(i));

        return temp;
    }

    // ========================================================================
    //  工具方法
    // ========================================================================
    private Sheet resolveSheet(Workbook wb, String sheetName) {
        Sheet s = (sheetName != null && !sheetName.isEmpty()) ? wb.getSheet(sheetName) : null;
        return s != null ? s : wb.getSheetAt(0);
    }

    private int resolveSplitCol(Sheet sheet, String splitColName) {
        Row headerRow = sheet.getRow(1);
        if (headerRow == null) throw new IllegalArgumentException("第2行表头不存在");
        int idx = findColumnIndex(headerRow, splitColName);
        if (idx < 0) throw new IllegalArgumentException("未找到列名: " + splitColName);
        return idx;
    }

    private String getSplitValue(Row row, int r, int splitColIdx, Map<String, String> mergeValues) {
        if (row == null) return "";
        String key = r + "_" + splitColIdx;
        return mergeValues.getOrDefault(key, getCellStringValue(row.getCell(splitColIdx)).trim());
    }

    private String extractBaseName(String fileName) {
        if (fileName == null || fileName.isEmpty()) return "拆分结果";
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    private int computeMaxCol(Sheet sheet, int maxRow) {
        int mc = 0;
        for (int r = 0; r <= maxRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > mc) mc = row.getLastCellNum();
        }
        return mc;
    }

    private Map<String, String> buildMergeValueMap(Sheet sheet) {
        Map<String, String> map = new LinkedHashMap<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            Row firstRow = sheet.getRow(region.getFirstRow());
            if (firstRow == null) continue;
            String val = getCellStringValue(firstRow.getCell(region.getFirstColumn())).trim();
            for (int mr = region.getFirstRow(); mr <= region.getLastRow(); mr++)
                for (int mc = region.getFirstColumn(); mc <= region.getLastColumn(); mc++)
                    map.put(mr + "_" + mc, val);
        }
        return map;
    }

    private boolean isFooterRow(Row row) {
        if (row == null) return false;
        for (int c = 0; c < row.getLastCellNum(); c++) {
            String val = getCellStringValue(row.getCell(c)).trim();
            for (String kw : FOOTER_KEYWORDS) if (val.contains(kw)) return true;
        }
        return false;
    }

    private int findColumnIndex(Row headerRow, String colName) {
        String target = colName.trim();
        for (int c = 0; c < headerRow.getLastCellNum(); c++)
            if (getCellStringValue(headerRow.getCell(c)).trim().equals(target)) return c;
        return -1;
    }

    private int detectFooterStart(Sheet sheet, int splitColIdx) {
        int maxRow = sheet.getLastRowNum();
        // 从后往前找，找到包含 SUM 公式的行即为表尾开始
        for (int r = maxRow; r >= 2; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() == CellType.FORMULA) {
                    String formula = cell.getCellFormula();
                    // 检测到 SUM 公式视为表尾行
                    if (formula != null && formula.toUpperCase().startsWith("SUM(")) {
                        return r;
                    }
                }
            }
        }
        return maxRow + 1;
    }

    private void copyCellValue(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING:  dst.setCellValue(src.getStringCellValue()); break;
            case NUMERIC: dst.setCellValue(src.getNumericCellValue()); break;
            case BOOLEAN: dst.setCellValue(src.getBooleanCellValue()); break;
            case FORMULA: dst.setCellFormula(src.getCellFormula()); break;
            default:      dst.setBlank();
        }
    }

    private String getSafeSheetName(Workbook wb, String baseName) {
        String name = baseName;
        int idx = 1;
        while (wb.getSheet(name) != null) name = baseName + "_" + idx++;
        return name;
    }

    private String sanitizeFilename(String name) {
        if (name == null || name.isEmpty()) return "unknown";
        name = name.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').trim();
        for (char ch : "\\/:*?\"<>|".toCharArray()) name = name.replace(String.valueOf(ch), "_");
        if (name.length() > 31) name = name.substring(0, 31);
        return name;
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:  return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try { return cell.getStringCellValue(); } catch (Exception e) { return cell.getCellFormula(); }
            default: return "";
        }
    }
}
