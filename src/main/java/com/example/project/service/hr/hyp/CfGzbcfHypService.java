package com.example.project.service.hr.hyp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * 对应 Python: cf_gzbcf_hyp.py（增强版）
 * 功能：按指定列拆分工资表，保留格式、公式、合并单元格，
 *       支持 公积金公式回填、社保条件填充、个税公式填充。
 */
@Service
public class CfGzbcfHypService {

    private static final Set<String> FOOTER_KEYWORDS = Set.of("合计", "总计", "汇总", "小计", "平均");
    private static final List<String> SHEETS_TO_REMOVE = List.of("各公司支出", "未发工资的社保扣除人员");

    // ========================================================================
    //  公开 API
    // ========================================================================

    /**
     * 获取 Excel 文件的所有工作表名称
     */
    public List<String> getSheetNames(MultipartFile file) throws Exception {
        List<String> sheets = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                sheets.add(wb.getSheetName(i));
            }
        }
        return sheets;
    }

    /**
     * 获取指定工作表指定行的表头列表（headerRow 为 1-based，如第2行传入2）
     */
    public List<String> getHeaders(MultipartFile file, String sheetName, int headerRow) throws Exception {
        List<String> headers = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                if (wb.getNumberOfSheets() > 0) sheet = wb.getSheetAt(0);
                else return headers;
            }
            Row row = sheet.getRow(headerRow - 1);
            if (row == null) return headers;
            int lastCell = row.getLastCellNum();
            for (int c = 0; c < lastCell; c++) {
                String val = getCellStringValue(row.getCell(c)).trim();
                if (!val.isEmpty()) headers.add(val);
            }
        }
        return headers;
    }

    /**
     * 获取指定列的所有唯一值（去重，保持顺序，自动处理合并单元格和表尾）
     */
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
                    if (mergeValues.containsKey(mergeKey)) {
                        val = mergeValues.get(mergeKey);
                    } else {
                        val = getCellStringValue(row.getCell(colIdx)).trim();
                    }
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
    //  核心拆分方法
    // ========================================================================

    /**
     * 按自定义分组拆分工资表
     */
    public String splitSalaryWithGroups(MultipartFile file, String originalName, String sheetName,
                                         String splitColName, List<List<String>> groups, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (splitColName == null || splitColName.trim().isEmpty()) throw new IllegalArgumentException("拆分依据列名不能为空");
        if (groups == null || groups.isEmpty()) return splitSalary(file, originalName, sheetName, splitColName, outputDir);

        File inputFile = new File(outputDir, "input_" + System.currentTimeMillis() + ".xlsx");
        file.transferTo(inputFile);

        try (Workbook wb = WorkbookFactory.create(inputFile)) {
            Sheet sheet = (sheetName != null && !sheetName.isEmpty()) ? wb.getSheet(sheetName) : null;
            if (sheet == null) sheet = wb.getSheetAt(0);
            String originalSheetName = sheet.getSheetName();

            Row headerRow = sheet.getRow(1);
            if (headerRow == null) throw new IllegalArgumentException("第2行表头不存在");
            int splitColIdx = findColumnIndex(headerRow, splitColName);
            if (splitColIdx < 0) throw new IllegalArgumentException("未找到列名: " + splitColName);

            int footerStart = detectFooterStart(sheet, splitColIdx);
            int maxRow = sheet.getLastRowNum();
            if (footerStart > maxRow) footerStart = maxRow + 1;

            int maxCol = computeMaxCol(sheet, maxRow);
            Map<String, String> mergeValues = buildMergeValueMap(sheet);

            // 构建值到组的映射
            Map<String, Integer> valueToGroup = new LinkedHashMap<>();
            Map<Integer, String> groupNames = new LinkedHashMap<>();
            Set<String> allGroupedValues = new LinkedHashSet<>();
            for (int gi = 0; gi < groups.size(); gi++) {
                List<String> groupValues = groups.get(gi);
                if (groupValues.isEmpty()) continue;
                groupNames.put(gi, sanitizeFilename(String.join("+", groupValues)));
                for (String v : groupValues) {
                    String trimmed = v.trim();
                    valueToGroup.put(trimmed, gi);
                    allGroupedValues.add(trimmed);
                }
            }

            // 按组收集行号
            Map<Integer, List<Integer>> groupRows = new LinkedHashMap<>();
            for (int i = 0; i < groups.size(); i++) groupRows.put(i, new ArrayList<>());
            Map<String, List<Integer>> ungrouped = new LinkedHashMap<>();
            String lastVal = "";

            for (int r = 2; r < footerStart; r++) {
                Row row = sheet.getRow(r);
                String val = getSplitValue(row, r, splitColIdx, mergeValues);
                if (val.isEmpty() && !lastVal.isEmpty()) val = lastVal;
                if (!val.isEmpty() && !isFooterRow(row, FOOTER_KEYWORDS)) {
                    Integer groupIdx = valueToGroup.get(val);
                    if (groupIdx != null) {
                        groupRows.get(groupIdx).add(r);
                    } else {
                        ungrouped.computeIfAbsent(val, k -> new ArrayList<>()).add(r);
                    }
                    lastVal = val;
                }
            }

            String zipName = extractBaseName(originalName) + "_拆分结果.zip";
            File zipFile = new File(outputDir, zipName);
            try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
                for (Map.Entry<Integer, List<Integer>> entry : groupRows.entrySet()) {
                    List<Integer> rows = entry.getValue();
                    if (rows.isEmpty()) continue;
                    String groupFileName = groupNames.get(entry.getKey());
                    generateAndZip(inputFile, originalSheetName, groupFileName, rows,
                            footerStart, maxRow, maxCol, splitColIdx, zos, outputDir);
                }
                for (Map.Entry<String, List<Integer>> entry : ungrouped.entrySet()) {
                    generateAndZip(inputFile, originalSheetName, entry.getKey(), entry.getValue(),
                            footerStart, maxRow, maxCol, splitColIdx, zos, outputDir);
                }
            }
            inputFile.delete();
            return zipFile.getAbsolutePath();
        }
    }

    /**
     * 按列值各自拆分工资表
     */
    public String splitSalary(MultipartFile file, String originalName, String sheetName,
                               String splitColName, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (splitColName == null || splitColName.trim().isEmpty()) throw new IllegalArgumentException("拆分依据列名不能为空");

        File inputFile = new File(outputDir, "input_" + System.currentTimeMillis() + ".xlsx");
        file.transferTo(inputFile);

        try (Workbook wb = WorkbookFactory.create(inputFile)) {
            Sheet sheet = (sheetName != null && !sheetName.isEmpty()) ? wb.getSheet(sheetName) : null;
            if (sheet == null) sheet = wb.getSheetAt(0);
            String originalSheetName = sheet.getSheetName();

            Row headerRow = sheet.getRow(1);
            if (headerRow == null) throw new IllegalArgumentException("第2行表头不存在");
            int splitColIdx = findColumnIndex(headerRow, splitColName);
            if (splitColIdx < 0) throw new IllegalArgumentException("未找到列名: " + splitColName);

            int footerStart = detectFooterStart(sheet, splitColIdx);
            int maxRow = sheet.getLastRowNum();
            if (footerStart > maxRow) footerStart = maxRow + 1;

            int maxCol = computeMaxCol(sheet, maxRow);
            Map<String, String> mergeValues = buildMergeValueMap(sheet);

            Map<String, List<Integer>> groups = new LinkedHashMap<>();
            String lastVal = "";
            for (int r = 2; r < footerStart; r++) {
                Row row = sheet.getRow(r);
                String val = getSplitValue(row, r, splitColIdx, mergeValues);
                if (val.isEmpty() && !lastVal.isEmpty()) val = lastVal;
                if (!val.isEmpty() && !isFooterRow(row, FOOTER_KEYWORDS)) {
                    groups.computeIfAbsent(val, k -> new ArrayList<>()).add(r);
                    lastVal = val;
                }
            }
            if (groups.isEmpty()) throw new IllegalArgumentException("未找到有效的拆分数据");

            String zipName = extractBaseName(originalName) + "_拆分结果.zip";
            File zipFile = new File(outputDir, zipName);
            try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
                for (Map.Entry<String, List<Integer>> entry : groups.entrySet()) {
                    generateAndZip(inputFile, originalSheetName, entry.getKey(), entry.getValue(),
                            footerStart, maxRow, maxCol, splitColIdx, zos, outputDir);
                }
            }
            return zipFile.getAbsolutePath();
        } finally {
            inputFile.delete();
        }
    }

    // ========================================================================
    //  辅助：生成单个文件并加入 ZIP
    // ========================================================================
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
    //  核心：生成单个部门的输出文件
    // ========================================================================
    private void processDepartmentFile(File file, String originalSheetName, String deptName,
                                        List<Integer> dataRows, int footerStart, int maxRow,
                                        int maxCol, int splitColIdx) throws Exception {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet oldSheet = wb.getSheet(originalSheetName);
            if (oldSheet == null) oldSheet = wb.getSheetAt(0);

            // ---- 步骤1: 过滤幽灵行（排除拆分列外全空的行） ----
            List<Integer> filteredRows = filterGhostRows(oldSheet, dataRows, maxCol, splitColIdx);
            if (filteredRows.isEmpty())
                throw new IllegalStateException("过滤后无有效数据行: " + deptName);

            // ---- 步骤2: 构建保留行列表与行号映射 ----
            List<Integer> keepRows = new ArrayList<>();
            keepRows.add(0);  // 第1行：大标题
            keepRows.add(1);  // 第2行：表头
            keepRows.addAll(filteredRows);

            // 添加原表尾行（从 footerStart 到文件末尾，保留 SUM 公式等）
            int sheetMaxRow = Math.max(maxRow, oldSheet.getLastRowNum());
            List<Integer> footerRows = new ArrayList<>();
            for (int r = footerStart; r <= sheetMaxRow; r++) {
                Row row = oldSheet.getRow(r);
                if (row == null) continue;
                boolean hasContent = false;
                for (int c = 0; c < maxCol; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && !getCellStringValue(cell).trim().isEmpty()) {
                        hasContent = true;
                        break;
                    }
                }
                if (hasContent) footerRows.add(r);
            }
            keepRows.addAll(footerRows);

            Set<Integer> keepSet = new HashSet<>(keepRows);
            Map<Integer, Integer> rowMap = new LinkedHashMap<>();
            int newR = 0;
            for (int oldR : keepRows) {
                rowMap.put(oldR, newR++);
            }
            int newDataLastRow = rowMap.get(filteredRows.get(filteredRows.size() - 1));
            int oldLastDataRow = footerStart - 1;
            int oldFirstDataRow = filteredRows.get(0);
            int newFirstDataRow = rowMap.get(oldFirstDataRow);
            int actualMaxRow = keepRows.size();  // 新文件最大行数（0-based）

            // ---- 步骤3: 扫描原表公积金公式（在创建新Sheet前完成） ----
            GjjScanResult gjjResult = scanGjjFormulas(oldSheet, maxCol, footerStart, splitColIdx);

            // ---- 步骤4: 创建新Sheet并逐行复制 ----
            String safeSheetName = getSafeSheetName(wb, sanitizeFilename(deptName));
            Sheet newSheet = wb.createSheet(safeSheetName);
            wb.setSheetOrder(safeSheetName, 0);

            for (int oldR : keepRows) {
                Row oldRow = oldSheet.getRow(oldR);
                Row newRow = newSheet.createRow(rowMap.get(oldR));
                if (oldRow == null) continue;

                for (int c = 0; c < maxCol; c++) {
                    Cell oldCell = oldRow.getCell(c);
                    if (oldCell == null) continue;
                    Cell newCell = newRow.createCell(c);

                    // 公式单元格：保留公式文本
                    if (oldCell.getCellType() == CellType.FORMULA) {
                        newCell.setCellFormula(oldCell.getCellFormula());
                    } else {
                        copyCellValue(oldCell, newCell);
                    }
                    copyCellStyle(oldCell, newCell);
                }
                newRow.setHeight(oldRow.getHeight());
            }

            // ---- 步骤5: 复制列宽 ----
            for (int c = 0; c < maxCol; c++) {
                int width = oldSheet.getColumnWidth(c);
                newSheet.setColumnWidth(c, width);
            }

            // ---- 步骤6: 重建合并单元格 ----
            int mergedCount = 0;
            for (int i = 0; i < oldSheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = oldSheet.getMergedRegion(i);
                Set<Integer> regionRows = new HashSet<>();
                for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) regionRows.add(r);
                if (keepSet.containsAll(regionRows)) {
                    newSheet.addMergedRegion(new CellRangeAddress(
                            rowMap.get(region.getFirstRow()), rowMap.get(region.getLastRow()),
                            region.getFirstColumn(), region.getLastColumn()));
                    mergedCount++;
                }
            }

            // ---- 步骤7: 调整所有已有公式的行号 ----
            adjustAllFormulas(newSheet, maxCol, actualMaxRow, rowMap,
                    oldLastDataRow, newDataLastRow, oldFirstDataRow, newFirstDataRow);

            // ---- 步骤8: 构建列名映射 ----
            Map<String, List<Integer>> colNameMap = new LinkedHashMap<>();
            Row newHeaderRow = newSheet.getRow(1);
            if (newHeaderRow != null) {
                for (int c = 0; c < maxCol; c++) {
                    String name = getCellStringValue(newHeaderRow.getCell(c)).trim();
                    if (!name.isEmpty()) {
                        colNameMap.computeIfAbsent(name, k -> new ArrayList<>()).add(c);
                    }
                }
            }

            // ---- 步骤9: 填充个税公式（覆盖所有数据行） ----
            fillTaxFormula(newSheet, colNameMap, newDataLastRow);

            // ---- 步骤10: 社保公式条件填充（检测个人扣款列，有数据才填） ----
            fillSocialSecurityFormulas(newSheet, colNameMap, newDataLastRow, oldMidColIndex(oldSheet, maxCol));

            // ---- 步骤11: 公积金公式回填 ----
            backfillGjjFormulas(newSheet, rowMap, gjjResult,
                    oldLastDataRow, newDataLastRow, actualMaxRow,
                    oldFirstDataRow, newFirstDataRow);

            // ---- 步骤12: 清理并保存 ----
            int oldIdx = wb.getSheetIndex(oldSheet);
            wb.removeSheetAt(oldIdx);

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
    //  幽灵行过滤
    // ========================================================================
    private List<Integer> filterGhostRows(Sheet sheet, List<Integer> dataRows, int maxCol, int splitColIdx) {
        List<Integer> result = new ArrayList<>();
        for (int oldR : dataRows) {
            Row row = sheet.getRow(oldR);
            if (row == null) continue;
            boolean hasContent = false;
            for (int c = 0; c < maxCol; c++) {
                if (c == splitColIdx) continue;
                Cell cell = row.getCell(c);
                if (cell != null && !getCellStringValue(cell).trim().isEmpty()) {
                    hasContent = true;
                    break;
                }
            }
            if (hasContent) result.add(oldR);
        }
        return result;
    }

    // ========================================================================
    //  公积金公式扫描
    // ========================================================================
    private static class GjjRecord {
        final int colIdx;
        final int oldRow;
        final String formula;
        GjjRecord(int colIdx, int oldRow, String formula) {
            this.colIdx = colIdx;
            this.oldRow = oldRow;
            this.formula = formula;
        }
    }

    private static class GjjScanResult {
        final List<GjjRecord> records = new ArrayList<>();
    }

    private GjjScanResult scanGjjFormulas(Sheet sheet, int maxCol, int footerStart, int splitColIdx) {
        GjjScanResult result = new GjjScanResult();
        Row headerRow = sheet.getRow(1);
        if (headerRow == null) return result;

        for (int c = 0; c < maxCol; c++) {
            String name = getCellStringValue(headerRow.getCell(c)).trim();
            if (!"公积金扣款-单位".equals(name) && !"公积金扣款-个人".equals(name)) continue;

            for (int r = 2; r < footerStart; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(c);
                if (cell == null) continue;

                String formula = null;
                if (cell.getCellType() == CellType.FORMULA) {
                    formula = cell.getCellFormula();
                }
                if (formula != null && !formula.isEmpty()) {
                    result.records.add(new GjjRecord(c, r, formula));
                }
            }
        }
        return result;
    }

    // ========================================================================
    //  个税公式填充
    //  对齐 Python：按表头名动态定位依赖列（替代硬编码 AK/AJ/AT/AM/AN），
    //  缺失时保留原表"个税"列数值并写日志，不覆盖。
    // ========================================================================

    /** 个税公式依赖字段 -> 表头名（与 Python tax_field_headers 对齐） */
    private static final Map<String, String> TAX_FIELD_HEADERS = Map.of(
            "taxable_base",        "税前工资",
            "subtract_deduction",  "专项附加扣除",
            "subtract_allowance",  "通讯费",
            "add_social",          "社保个人扣款",
            "add_fund",            "公积金扣款-个人"
    );

    private void fillTaxFormula(Sheet newSheet, Map<String, List<Integer>> colNameMap, int newDataLastRow) {
        List<Integer> taxCols = colNameMap.get("个税");
        if (taxCols == null || taxCols.isEmpty()) return;
        int taxCol = taxCols.get(0);

        // 按表头名动态定位各依赖列
        Map<String, Integer> taxFieldCols = new HashMap<>();
        List<String> taxMissing = new ArrayList<>();
        for (Map.Entry<String, String> entry : TAX_FIELD_HEADERS.entrySet()) {
            List<Integer> cols = colNameMap.get(entry.getValue());
            if (cols != null && !cols.isEmpty()) {
                taxFieldCols.put(entry.getKey(), cols.get(0));
            } else {
                taxMissing.add(entry.getValue());
            }
        }
        if (!taxMissing.isEmpty()) {
            System.out.println("[WARN] 个税依赖列未全部定位（缺失 " + taxMissing + "），保留原表'个税'列数值，未覆盖写入公式");
            return;
        }

        String tb = getColumnLetter(taxFieldCols.get("taxable_base"));        // 税前工资
        String sd = getColumnLetter(taxFieldCols.get("subtract_deduction"));  // 专项附加扣除
        String ea = getColumnLetter(taxFieldCols.get("subtract_allowance"));  // 通讯费
        String sp = getColumnLetter(taxFieldCols.get("add_social"));          // 社保个人扣款
        String fp = getColumnLetter(taxFieldCols.get("add_fund"));            // 公积金扣款-个人

        for (int r = 2; r <= newDataLastRow; r++) {
            Row row = newSheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(taxCol);
            if (cell == null) cell = row.createCell(taxCol);
            int excelRow = r + 1;  // r 是 0-based，Excel 行号 = r+1
            cell.setCellFormula(
                "-ROUND(MAX((" + tb + excelRow
                + "-5000-" + sd + excelRow
                + "-" + ea + excelRow
                + "+" + sp + excelRow
                + "+" + fp + excelRow
                + ")*{3,10,20,25,30,35,45}%-5*{0,42,282,532,882,1432,3032},),2)"
            );
        }
    }

    // ========================================================================
    //  社保公式条件填充
    //  对齐 Python：动态定位"身份证编号"列（替代硬编码 E 列），找不到时回退 E 列。
    // ========================================================================
    private void fillSocialSecurityFormulas(Sheet newSheet, Map<String, List<Integer>> colNameMap,
                                             int newDataLastRow, int midCol) {
        List<Integer> ssPersonalCols = colNameMap.get("社保个人扣款");
        List<Integer> ssCompanyCols = colNameMap.get("社保单位扣款");
        if (ssPersonalCols == null || ssPersonalCols.isEmpty()) return;
        if (ssCompanyCols == null || ssCompanyCols.isEmpty()) return;

        int ssPersonalCol = ssPersonalCols.get(0);
        int ssCompanyCol = ssCompanyCols.get(ssCompanyCols.size() - 1);

        // 动态定位"身份证编号"列；缺失时回退 E 列
        String idLetter;
        List<Integer> idCols = colNameMap.get("身份证编号");
        if (idCols != null && !idCols.isEmpty()) {
            idLetter = getColumnLetter(idCols.get(0));
        } else {
            idLetter = "E";
            System.out.println("[WARN] 未找到'身份证编号'列，社保 XLOOKUP 查找键回退为 E 列");
        }

        for (int r = 2; r <= newDataLastRow; r++) {
            Row row = newSheet.getRow(r);
            if (row == null) continue;

            Cell personalCell = row.getCell(ssPersonalCol);
            boolean hasData = personalCell != null && !getCellStringValue(personalCell).trim().isEmpty();

            if (!hasData) {
                // 检查是否为数值（包括0）
                if (personalCell != null && personalCell.getCellType() == CellType.NUMERIC) {
                    hasData = true;
                }
            }

            if (hasData) {
                int excelRow = r + 1;
                // 社保单位扣款（"社保单位扣款"列中最右一列）
                Cell companyCell = row.getCell(ssCompanyCol);
                if (companyCell == null) companyCell = row.createCell(ssCompanyCol);
                companyCell.setCellFormula(
                    "XLOOKUP(" + idLetter + excelRow + ",社保!$A$2:$A$110,-社保!$C$2:$C$110,0)");

                // 社保个人扣款（覆盖原值）
                Cell pCell = row.getCell(ssPersonalCol);
                if (pCell == null) pCell = row.createCell(ssPersonalCol);
                pCell.setCellFormula(
                    "XLOOKUP(" + idLetter + excelRow + ",社保!$A$2:$A$110,-社保!$D$2:$D$110,0)");
            }
        }
    }

    /**
     * 获取原表的中间列索引（当找不到"出勤率"列时回退使用）
     */
    private int oldMidColIndex(Sheet oldSheet, int maxCol) {
        Row headerRow = oldSheet.getRow(1);
        if (headerRow != null) {
            for (int c = 0; c < maxCol; c++) {
                if (getCellStringValue(headerRow.getCell(c)).contains("出勤率")) return c;
            }
        }
        return Math.min(maxCol / 2, maxCol);
    }

    // ========================================================================
    //  公积金公式回填
    // ========================================================================
    private void backfillGjjFormulas(Sheet newSheet, Map<Integer, Integer> rowMap, GjjScanResult gjjResult,
                                      int oldLastDataRow, int newDataLastRow, int newMaxRow,
                                      int oldFirstDataRow, int newFirstDataRow) {
        int filled = 0;
        for (GjjRecord rec : gjjResult.records) {
            if (!rowMap.containsKey(rec.oldRow)) continue;
            int newR = rowMap.get(rec.oldRow);
            String adjusted = adjustFormula(rec.formula, rowMap,
                    oldLastDataRow, newDataLastRow, newMaxRow,
                    oldFirstDataRow, newFirstDataRow);
            Row row = newSheet.getRow(newR);
            Cell cell = row.getCell(rec.colIdx);
            if (cell == null) cell = row.createCell(rec.colIdx);
            cell.setCellFormula(adjusted);
            filled++;
        }
    }

    // ========================================================================
    //  调整所有公式行号（扫描新Sheet中所有公式单元格）
    // ========================================================================
    private void adjustAllFormulas(Sheet sheet, int maxCol, int newMaxRow,
                                    Map<Integer, Integer> rowMap,
                                    int oldLastDataRow, int newDataLastRow,
                                    int oldFirstDataRow, int newFirstDataRow) {
        for (int r = 0; r <= newMaxRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < maxCol; c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() == CellType.FORMULA) {
                    String oldFormula = cell.getCellFormula();
                    String newFormula = adjustFormula(oldFormula, rowMap,
                            oldLastDataRow, newDataLastRow, newMaxRow,
                            oldFirstDataRow, newFirstDataRow);
                    if (!oldFormula.equals(newFormula)) {
                        cell.setCellFormula(newFormula);
                    }
                }
            }
        }
    }

    // ========================================================================
    //  动态生成总计行
    // ========================================================================
    private void generateSummaryRow(Sheet newSheet, int newDataLastRow, int maxCol, Workbook wb) {
        Row totalRow = newSheet.createRow(newDataLastRow + 1);
        int sumStartCol = findAttendanceCol(newSheet, maxCol);

        CellStyle totalStyle = createSummaryStyle(wb);
        for (int c = 0; c < maxCol; c++) {
            Cell cell = totalRow.createCell(c);
            cell.setCellStyle(totalStyle);
            if (c < sumStartCol) {
                if (c == 0) cell.setCellValue("合计");
            } else {
                String colLetter = getColumnLetter(c);
                cell.setCellFormula("SUM(" + colLetter + "3:" + colLetter + (newDataLastRow + 1) + ")");
            }
        }
    }

    private int findAttendanceCol(Sheet sheet, int maxCol) {
        Row headerRow = sheet.getRow(1);
        if (headerRow != null) {
            for (int c = 0; c < maxCol; c++) {
                if (getCellStringValue(headerRow.getCell(c)).contains("出勤率")) return c;
            }
        }
        return Math.min(maxCol / 2, maxCol);
    }

    private CellStyle createSummaryStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        if (wb instanceof XSSFWorkbook) {
            style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 217, (byte) 217, (byte) 217}, null));
        } else {
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        }
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    // ========================================================================
    //  公式行号调整（增强版，分段处理跨Sheet引用）
    // ========================================================================
    /**
     * 调整公式中的行号引用。
     * 采用分段策略：将跨Sheet引用保护起来，只调整本Sheet部分的单元格行号，
     * 最后还原跨Sheet引用。
     *
     * @param oldFirstDataRow  原文件中本部门第一条数据行号
     * @param newFirstDataRow  新文件中本部门第一条数据行号（用于跨部门下界映射）
     * @param newMaxRow        新文件最大行号（超出原数据区的行映射到此）
     */
    private String adjustFormula(String formula, Map<Integer, Integer> rowMap,
                                  int oldLastDataRow, int newDataLastRow, int newMaxRow,
                                  int oldFirstDataRow, int newFirstDataRow) {
        if (formula == null || formula.isEmpty()) return formula;

        // ---- 阶段1: 匹配跨Sheet引用，切分为段落 ----
        Pattern crossPattern = Pattern.compile(
                "('[^']+'|[A-Za-z0-9_\\u4e00-\\u9fa5]+)!" +
                "(\\$?[A-Za-z]{1,3}\\$?\\d+" +
                "(:\\$?[A-Za-z]{1,3}\\$?\\d+)?)");

        List<int[]> crossSpans = new ArrayList<>();  // [start, end]
        Matcher crossMatcher = crossPattern.matcher(formula);
        while (crossMatcher.find()) {
            crossSpans.add(new int[]{crossMatcher.start(), crossMatcher.end()});
        }

        // 切分段落
        List<Segment> segments = new ArrayList<>();
        int pos = 0;
        for (int[] span : crossSpans) {
            if (pos < span[0]) segments.add(new Segment(false, formula.substring(pos, span[0])));
            segments.add(new Segment(true, formula.substring(span[0], span[1])));
            pos = span[1];
        }
        if (pos < formula.length()) segments.add(new Segment(false, formula.substring(pos)));

        // ---- 阶段2: 只对本Sheet段做行号替换 ----
        Pattern localPattern = Pattern.compile("(\\$?[A-Za-z]{1,3}\\$?)(\\d+)");
        StringBuilder result = new StringBuilder();

        for (Segment seg : segments) {
            if (seg.isCross) {
                result.append(seg.text);
            } else {
                Matcher m = localPattern.matcher(seg.text);
                StringBuffer sb = new StringBuffer();
                while (m.find()) {
                    String colPart = m.group(1);
                    int rowNum = Integer.parseInt(m.group(2));  // Excel 1-based
                    int newNum = mapRowNumber(rowNum, rowMap,
                            oldLastDataRow, newDataLastRow, newMaxRow,
                            oldFirstDataRow, newFirstDataRow);
                    m.appendReplacement(sb, Matcher.quoteReplacement(colPart + newNum));
                }
                m.appendTail(sb);
                result.append(sb);
            }
        }
        return result.toString();
    }

    private int mapRowNumber(int rowNum, Map<Integer, Integer> rowMap,
                              int oldLastDataRow, int newDataLastRow, int newMaxRow,
                              int oldFirstDataRow, int newFirstDataRow) {
        int zeroBased = rowNum - 1;  // 转为0-based

        // 在行映射表中
        if (rowMap.containsKey(zeroBased)) {
            return rowMap.get(zeroBased) + 1;
        }
        // 原数据区最后一行 → 映射到新数据区最后一行
        if (rowNum == oldLastDataRow + 1) {
            return newDataLastRow + 1;
        }
        // 超出原数据区 → 映射到新文件最大行
        if (rowNum > oldLastDataRow + 1) {
            return newMaxRow + 1;
        }
        // 低于本部门第一条数据行 → 映射到新文件第一条数据行
        if (oldFirstDataRow > 0 && zeroBased < oldFirstDataRow) {
            return newFirstDataRow + 1;
        }
        // 在原数据范围内但不在rowMap中（其他部门行）→ 映射到新文件第一条数据行
        if (zeroBased >= 2 && zeroBased <= oldLastDataRow && !rowMap.containsKey(zeroBased)) {
            return newFirstDataRow + 1;
        }
        return rowNum;
    }

    private static class Segment {
        final boolean isCross;
        final String text;
        Segment(boolean isCross, String text) { this.isCross = isCross; this.text = text; }
    }

    // ========================================================================
    //  工具方法
    // ========================================================================

    private String getSplitValue(Row row, int r, int splitColIdx, Map<String, String> mergeValues) {
        if (row == null) return "";
        String mergeKey = r + "_" + splitColIdx;
        if (mergeValues.containsKey(mergeKey)) return mergeValues.get(mergeKey);
        return getCellStringValue(row.getCell(splitColIdx)).trim();
    }

    private String extractBaseName(String fileName) {
        if (fileName == null || fileName.isEmpty()) return "拆分结果";
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    private int computeMaxCol(Sheet sheet, int maxRow) {
        int maxCol = 0;
        for (int r = 0; r <= maxRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > maxCol) maxCol = row.getLastCellNum();
        }
        return maxCol;
    }

    private Map<String, String> buildMergeValueMap(Sheet sheet) {
        Map<String, String> map = new LinkedHashMap<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            Row firstRow = sheet.getRow(region.getFirstRow());
            if (firstRow != null) {
                String val = getCellStringValue(firstRow.getCell(region.getFirstColumn())).trim();
                for (int mr = region.getFirstRow(); mr <= region.getLastRow(); mr++) {
                    for (int mc = region.getFirstColumn(); mc <= region.getLastColumn(); mc++) {
                        map.put(mr + "_" + mc, val);
                    }
                }
            }
        }
        return map;
    }

    private boolean isFooterRow(Row row, Set<String> keywords) {
        if (row == null) return false;
        for (int c = 0; c < row.getLastCellNum(); c++) {
            String val = getCellStringValue(row.getCell(c)).trim();
            for (String kw : keywords) {
                if (val.contains(kw)) return true;
            }
        }
        return false;
    }

    private int findColumnIndex(Row headerRow, String colName) {
        String target = colName.trim();
        for (int c = 0; c < headerRow.getLastCellNum(); c++) {
            if (getCellStringValue(headerRow.getCell(c)).trim().equals(target)) return c;
        }
        return -1;
    }

    private int detectFooterStart(Sheet sheet, int splitColIdx) {
        int maxRow = sheet.getLastRowNum();
        for (int r = maxRow; r >= 2; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() == CellType.FORMULA) {
                    String formula = cell.getCellFormula();
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

    private void copyCellStyle(Cell src, Cell dst) {
        dst.setCellStyle(src.getCellStyle());
    }

    private String getSafeSheetName(Workbook wb, String baseName) {
        String name = baseName;
        int idx = 1;
        while (wb.getSheet(name) != null) {
            name = baseName + "_" + idx++;
        }
        return name;
    }

    private String sanitizeFilename(String name) {
        if (name == null || name.isEmpty()) return "unknown";
        name = name.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').trim();
        for (char ch : "\\/:*?\"<>|".toCharArray()) {
            name = name.replace(String.valueOf(ch), "_");
        }
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
            default:      return "";
        }
    }

    private String getColumnLetter(int colIndex) {
        if (colIndex < 0) return "A";
        StringBuilder sb = new StringBuilder();
        int num = colIndex + 1;
        while (num > 0) {
            int rem = (num - 1) % 26;
            sb.insert(0, (char) ('A' + rem));
            num = (num - 1) / 26;
        }
        return sb.toString();
    }
}
