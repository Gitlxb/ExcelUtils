package com.example.project.service.hr.hyp;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
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
 * 孝心金拆分服务（对应桌面版 hr/hyp/cf_xxj_hyp.py）。
 *
 * <p>表结构约定（与源表「26年第三季度孝心金.xlsx」一致）：
 * <ul>
 *   <li>第 1 行：标题（A1:H1 合并），输出时原文照抄</li>
 *   <li>第 2 行：表头，拆分依据列下拉的来源</li>
 *   <li>第 3 行起：数据区</li>
 *   <li>空行之后：表尾（汇总标签行 + 汇总公式行）</li>
 * </ul>
 *
 * <p>与「工资表拆分（简化版）」的关键差异：
 * <ol>
 *   <li><b>按合并块分段</b>而非按值分组：同名但不连续的两段分别输出为独立文件
 *       （例：源表「浙江锦途」在第 3-46 行与第 125-137 行 → 浙江锦途.xlsx、浙江锦途_2.xlsx）</li>
 *   <li>不含个税/社保/公积金公式回填，仅保留源表自身的 =SUM(...) 与 =H{r}/500</li>
 *   <li>表尾保留 <b>连续区间</b>（含分隔空行），使输出行数与源表结构一致</li>
 *   <li>输出 Sheet 名沿用源工作表名</li>
 *   <li>ZIP 内文件名重名自动加序号</li>
 * </ol>
 */
@Service
public class XxjSplitService {

    /** 0-based：表头区起始行（Excel 第 2 行） */
    private static final int HEADER_ROW = 1;
    // 数据起始行不再硬编码，由 detectDataStartRow 动态识别：单行表头 -> 2，三行表头 -> 4

    /** 表尾标签行常见词（避免被误判为数据行） */
    private static final Set<String> SUMMARY_KEYWORDS = Set.of(
            "合计", "总计", "汇总", "小计", "平均", "金额", "个数", "人数");

    /** 金额列候选表头名（用于分组预览的金额合计） */
    private static final List<String> AMOUNT_HEADER_CANDIDATES =
            List.of("金额", "金额(元)", "金额（元）", "补贴金额", "孝心金");

    private static final int MAX_SHEET_NAME_LEN = 31;
    private static final int MAX_FILENAME_LEN = 120;

    /** 一个拆分分组：拆分列取值相同的所有数据行（同名不连续时也并入同一组） */
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Segment {
        /** 组序号（0-based，按首次出现顺序，前端组合拆分用它引用分组） */
        public int index;
        /** 拆分列的值 */
        public String value;
        /** 首行（1-based，Excel 行号，便于界面展示） */
        public int startRow;
        /** 末行（1-based） */
        public int endRow;
        /** 数据行数 */
        public int count;
        /** 金额合计（无金额列时为 0） */
        public double amount;
        /** 行范围展示文本：单处为「第5-48行」，跨多处为「第5-48行 等2段」 */
        public String rangeText;

        /** 内部使用：该组包含的 0-based 行号（transient，不参与 JSON 序列化） */
        public transient List<Integer> rows = new ArrayList<>();

        public Segment() {
        }

        public int getIndex() { return index; }
        public String getValue() { return value; }
        public int getStartRow() { return startRow; }
        public int getEndRow() { return endRow; }
        public int getCount() { return count; }
        public double getAmount() { return amount; }
        public String getRangeText() { return rangeText; }
    }

    // ========================================================================
    //  公开 API
    // ========================================================================

    public List<String> getSheetNames(MultipartFile file) throws Exception {
        List<String> sheets = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) sheets.add(wb.getSheetName(i));
        }
        return sheets;
    }

    /**
     * 识别表头（兼容单行表头与多行表头两种格式）。
     *
     * <p>列名逐列在「表头区」内取最后一个非空、非公式的值：A~F 可能落在表头区首行
     * （靠跨行合并撑着），G/H 可能落在表头区末行；表头区里的「总金额」「个数」等
     * 汇总标签与其公式都不算列名。
     *
     * @param headerRow 兼容旧接口签名保留，实际范围由 detectDataStartRow 动态确定
     */
    public List<String> getHeaders(MultipartFile file, String sheetName, int headerRow) throws Exception {
        List<String> headers = new ArrayList<>();
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = resolveSheet(wb, sheetName);
            if (sheet == null) return headers;
            int dataStartRow = detectDataStartRow(sheet);
            headers.addAll(buildHeaderMap(sheet, dataStartRow).keySet());
        }
        return headers;
    }

    /** 识别拆分分组（供前端预览与组合拆分弹窗使用） */
    public List<Segment> getSegments(MultipartFile file, String sheetName, String colName) throws Exception {
        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = resolveSheet(wb, sheetName);
            if (sheet == null) return new ArrayList<>();
            int dataStartRow = detectDataStartRow(sheet);
            int splitColIdx = resolveSplitCol(sheet, colName, dataStartRow);
            return buildSegments(sheet, splitColIdx, dataStartRow);
        }
    }

    /** 直接拆分：每个段输出一个文件 */
    public String splitXxj(MultipartFile file, String originalName, String sheetName,
                           String splitColName, String outputDir) throws Exception {
        return doSplit(file, originalName, sheetName, splitColName, null, outputDir);
    }

    /** 组合拆分：组内段合并为一个文件，未入组的段各自输出 */
    public String splitXxjWithGroups(MultipartFile file, String originalName, String sheetName,
                                     String splitColName, List<List<Integer>> groups, String outputDir) throws Exception {
        return doSplit(file, originalName, sheetName, splitColName, groups, outputDir);
    }

    // ========================================================================
    //  核心拆分
    // ========================================================================

    private String doSplit(MultipartFile file, String originalName, String sheetName,
                           String splitColName, List<List<Integer>> groups, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (splitColName == null || splitColName.trim().isEmpty()) {
            throw new IllegalArgumentException("拆分依据列名不能为空");
        }

        File inputFile = new File(outputDir, "input_xxj_" + System.currentTimeMillis() + ".xlsx");
        file.transferTo(inputFile);

        try (Workbook wb = WorkbookFactory.create(inputFile)) {
            Sheet sheet = resolveSheet(wb, sheetName);
            if (sheet == null) throw new IllegalArgumentException("未找到工作表");
            String originalSheetName = sheet.getSheetName();
            int dataStartRow = detectDataStartRow(sheet);
            int splitColIdx = resolveSplitCol(sheet, splitColName, dataStartRow);
            int maxRow = sheet.getLastRowNum();
            int maxCol = computeMaxCol(sheet, maxRow);

            int[] footer = detectFooterRange(sheet, splitColIdx, dataStartRow);
            int footerStart = footer[0];
            int footerEnd = footer[1];

            // buildSegments 内部已统一算好 index / 行范围 / 行数 / 金额合计
            List<Segment> segments = buildSegments(sheet, splitColIdx, dataStartRow);
            if (segments.isEmpty()) throw new IllegalArgumentException("未找到有效的拆分数据");

            // 待处理任务：{输出文件名, 段列表}
            List<String[]> tasks = new ArrayList<>();
            Set<Integer> assigned = new HashSet<>();
            if (groups != null && !groups.isEmpty()) {
                for (List<Integer> grp : groups) {
                    List<Segment> members = new ArrayList<>();
                    for (Integer idx : grp) {
                        if (idx != null && idx >= 0 && idx < segments.size() && assigned.add(idx)) {
                            members.add(segments.get(idx));
                        }
                    }
                    if (!members.isEmpty()) tasks.add(new String[]{composeGroupName(members), joinRowIds(members)});
                }
            }
            for (Segment s : segments) {
                if (!assigned.contains(s.index)) {
                    tasks.add(new String[]{s.value, joinRowIds(Collections.singletonList(s))});
                }
            }

            String zipName = extractBaseName(originalName) + "_拆分结果.zip";
            File zipFile = new File(outputDir, zipName);
            Set<String> usedEntryNames = new LinkedHashSet<>();

            try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
                for (String[] task : tasks) {
                    String entryName = uniqueEntryName(sanitizeFilename(task[0]), usedEntryNames);
                    List<Integer> rowIds = parseRowIds(task[1]);
                    File tmpFile = new File(outputDir, "xxj_tmp_" + System.currentTimeMillis() + ".xlsx");
                    Files.copy(inputFile.toPath(), tmpFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
                    try {
                        processSegmentFile(tmpFile, originalSheetName, rowIds,
                                footerStart, footerEnd, maxCol, splitColIdx, dataStartRow, task[0]);
                        zos.putNextEntry(new ZipEntry(entryName));
                        Files.copy(tmpFile.toPath(), zos);
                        zos.closeEntry();
                    } finally {
                        tmpFile.delete();
                    }
                }
            }
            return zipFile.getAbsolutePath();
        } finally {
            inputFile.delete();
        }
    }

    /** 生成单个段（或组合）的输出文件，就地写回 tmpFile */
    private void processSegmentFile(File file, String originalSheetName, List<Integer> rowIds,
                                    int footerStart, int footerEnd, int maxCol, int splitColIdx,
                                    int dataStartRow, String label) throws Exception {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet oldSheet = wb.getSheet(originalSheetName);
            if (oldSheet == null) oldSheet = wb.getSheetAt(0);

            // 1. 幽灵行过滤（拆分列以外的列全空视为无效行）
            List<Integer> filteredRows = new ArrayList<>();
            for (int oldR : rowIds) {
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
            if (filteredRows.isEmpty()) throw new IllegalStateException("过滤后无有效数据行：" + label);
            filteredRows.sort(Integer::compareTo);

            // 2. 保留行：标题行 + 完整表头区（第 0 ~ dataStartRow-1 行，单行/多行表头都适用）
            //            + 数据行 + 表尾连续区间（含分隔空行）
            List<Integer> keepRows = new ArrayList<>();
            for (int r = 0; r < dataStartRow; r++) keepRows.add(r);
            keepRows.addAll(filteredRows);
            if (footerEnd >= footerStart) {
                for (int r = footerStart; r <= footerEnd; r++) keepRows.add(r);
            }
            Set<Integer> keepSet = new HashSet<>(keepRows);
            Map<Integer, Integer> rowMap = new LinkedHashMap<>();
            int newR = 0;
            for (int oldR : keepRows) rowMap.put(oldR, newR++);

            int newDataLastRow = rowMap.get(filteredRows.get(filteredRows.size() - 1));
            int oldLastDataRow = footerStart - 1;

            // 3. 新 Sheet（先用临时名，避免与待删除的旧表同名冲突）
            String tmpSheetName = getSafeSheetName(wb, "__xxj_tmp__");
            Sheet newSheet = wb.createSheet(tmpSheetName);
            wb.setSheetOrder(tmpSheetName, 0);

            // 4. 逐行复制：值 / 公式 / 样式 / 行高
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

            // 5. 列宽按源文件复制
            for (int c = 0; c < maxCol; c++) {
                newSheet.setColumnWidth(c, oldSheet.getColumnWidth(c));
            }

            // 6. 重建合并单元格（整块都在保留行内才重建）
            //    拆分列单独处理：该列在这一组的所有数据行上取值相同，源表里可能分成多块
            //    （同名值在源表中不连续时），这里统一合并成一个整块，避免出现两个同名单元格。
            for (int i = 0; i < oldSheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = oldSheet.getMergedRegion(i);
                // 只跳过「数据区、且仅覆盖拆分列」的合并块（交给下面整组统一合并）。
                // 判断必须同时限定单列 + 数据区：若只写首列/末列包含拆分列，
                // 会把 A1:H1 这种跨列标题合并也误跳过，导致输出文件标题塌进 A1。
                if (region.getFirstColumn() == splitColIdx
                        && region.getLastColumn() == splitColIdx
                        && region.getFirstRow() >= dataStartRow) {
                    continue;
                }
                Set<Integer> regionRows = new HashSet<>();
                for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) regionRows.add(r);
                if (keepSet.containsAll(regionRows)) {
                    newSheet.addMergedRegion(new CellRangeAddress(
                            rowMap.get(region.getFirstRow()), rowMap.get(region.getLastRow()),
                            region.getFirstColumn(), region.getLastColumn()));
                }
            }

            // 拆分列：整组数据行合并为一个单元格（行在输出中连续时才合并）
            int firstNewRow = rowMap.get(filteredRows.get(0));
            int lastNewRow = rowMap.get(filteredRows.get(filteredRows.size() - 1));
            if (lastNewRow > firstNewRow && (lastNewRow - firstNewRow + 1) == filteredRows.size()) {
                newSheet.addMergedRegion(new CellRangeAddress(firstNewRow, lastNewRow,
                        splitColIdx, splitColIdx));
            }

            // 7. 删除旧 Sheet（名字释放后可复用）
            wb.removeSheetAt(wb.getSheetIndex(oldSheet));

            // 8. 汇总公式行号重排：=SUM(H3:H137) -> =SUM(H3:H15)、=H140/500 -> =H18/500
            int newMaxRow = newSheet.getLastRowNum();
            for (int r = 0; r <= newMaxRow; r++) {
                Row row = newSheet.getRow(r);
                if (row == null) continue;
                for (int c = 0; c < maxCol; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && cell.getCellType() == CellType.FORMULA) {
                        String f = cell.getCellFormula();
                        String adjusted = adjustFormula(f, rowMap, oldLastDataRow, newDataLastRow, newMaxRow + 1);
                        if (!f.equals(adjusted)) cell.setCellFormula(adjusted);
                    }
                }
            }

            // 9. Sheet 名沿用源工作表名
            wb.setSheetName(wb.getSheetIndex(newSheet), sanitizeSheetName(originalSheetName));

            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
    }

    // ========================================================================
    //  分段（按合并块）
    // ========================================================================

    /**
     * 按拆分列的取值聚合分组。
     *
     * <p>同名值无论出现在几个不连续的位置，都并入同一个分组（输出为同一个文件），
     * 分组顺序按该值在源表中首次出现的顺序。
     */
    private List<Segment> buildSegments(Sheet sheet, int splitColIdx, int dataStartRow) {
        Map<String, String> mergeValues = buildMergeValueMap(sheet);
        int footerStart = detectFooterRange(sheet, splitColIdx, dataStartRow)[0];

        Map<String, Segment> groups = new LinkedHashMap<>();
        String lastVal = "";

        for (int r = dataStartRow; r < footerStart; r++) {
            Row row = sheet.getRow(r);
            String val = "";
            String key = r + "_" + splitColIdx;
            if (mergeValues.containsKey(key)) {
                val = mergeValues.get(key);
            } else if (row != null) {
                val = getCellStringValue(row.getCell(splitColIdx)).trim();
            }
            // 拆分列为空时沿用上一行的值（合并/空单元格容错）
            if (val.isEmpty()) val = lastVal;
            if (val.isEmpty()) continue;

            Segment seg = groups.get(val);
            if (seg == null) {
                seg = newSegment(val, r);
                groups.put(val, seg);
            } else {
                seg.rows.add(r);
            }
            lastVal = val;
        }

        List<Segment> segments = new ArrayList<>(groups.values());

        // 统一在此计算 index / 行范围 / 行数 / 金额合计，避免调用方各自实现导致漏算
        int amountColIdx = resolveAmountCol(sheet, dataStartRow);
        for (int i = 0; i < segments.size(); i++) {
            Segment s = segments.get(i);
            s.index = i;
            s.rows.sort(Integer::compareTo);
            s.startRow = s.rows.get(0) + 1;
            s.endRow = s.rows.get(s.rows.size() - 1) + 1;
            s.count = s.rows.size();
            s.amount = sumAmount(sheet, s.rows, amountColIdx);
            s.rangeText = buildRangeText(s.rows);
        }
        return segments;
    }

    /** 行范围文本：该组在源表中的首行 ~ 末行（同名值跨多处时也只显示一个范围） */
    private String buildRangeText(List<Integer> rows) {
        if (rows == null || rows.isEmpty()) return "";
        List<Integer> sorted = new ArrayList<>(rows);
        sorted.sort(Integer::compareTo);
        return "第" + (sorted.get(0) + 1) + "-" + (sorted.get(sorted.size() - 1) + 1) + "行";
    }

    private Segment newSegment(String value, int zeroBasedRow) {
        Segment s = new Segment();
        s.value = value;
        s.rows = new ArrayList<>();
        s.rows.add(zeroBasedRow);
        return s;
    }

    // ========================================================================
    //  表尾检测
    // ========================================================================

    /**
     * 返回 {表尾起始行, 表尾结束行}（0-based；结束行小于起始行表示无表尾）。
     *
     * <p>数据区最后一行 = 拆分列上最后一个「有效值」所在行，有效值需排除：
     * 空值、公式、汇总关键词、与表头同名的标签（如表尾的 H=金额 / I=个数）。
     * 表尾结束行取表尾区间内最后一个非空行，从而保留分隔空行。
     */
    private int[] detectFooterRange(Sheet sheet, int splitColIdx, int dataStartRow) {
        int maxRow = sheet.getLastRowNum();
        Map<String, String> mergeValues = buildMergeValueMap(sheet);
        Set<String> headerNames = buildHeaderMap(sheet, dataStartRow).keySet();

        int lastDataRow = dataStartRow - 1;
        for (int r = dataStartRow; r <= maxRow; r++) {
            Row row = sheet.getRow(r);
            String val;
            String key = r + "_" + splitColIdx;
            if (mergeValues.containsKey(key)) {
                val = mergeValues.get(key);
            } else if (row != null) {
                val = getCellStringValue(row.getCell(splitColIdx)).trim();
            } else {
                continue;
            }
            val = val == null ? "" : val.trim();
            if (val.isEmpty() || val.startsWith("=")) continue;
            if (SUMMARY_KEYWORDS.contains(val) || headerNames.contains(val)) continue;
            lastDataRow = r;
        }

        int footerStart = lastDataRow + 1;
        int footerEnd = footerStart - 1;
        int maxColAll = computeMaxCol(sheet, maxRow);
        for (int r = footerStart; r <= maxRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < maxColAll; c++) {
                Cell cell = row.getCell(c);
                if (cell != null && !getCellStringValue(cell).trim().isEmpty()) {
                    footerEnd = r;
                    break;
                }
            }
        }
        return new int[]{footerStart, footerEnd};
    }

    // ========================================================================
    //  公式行号重排（占位符保护跨 Sheet 引用）
    // ========================================================================

    private String adjustFormula(String formula, Map<Integer, Integer> rowMap,
                                 int oldLastDataRow, int newDataLastRow, int newMaxRow) {
        if (formula == null || formula.isEmpty()) return formula;

        Pattern crossPattern = Pattern.compile(
                "('[^']+'|[A-Za-z0-9_\\u4e00-\\u9fa5]+)!" +
                "(\\$?[A-Za-z]{1,3}\\$?\\d+(:\\$?[A-Za-z]{1,3}\\$?\\d+)?)");
        List<String> crossPh = new ArrayList<>();
        Matcher crossMatcher = crossPattern.matcher(formula);
        while (crossMatcher.find()) crossPh.add(crossMatcher.group(0));

        String temp = formula;
        for (int i = crossPh.size() - 1; i >= 0; i--) {
            temp = temp.replace(crossPh.get(i), "__CROSS" + i + "__");
        }

        Pattern localPattern = Pattern.compile("(\\$?[A-Za-z]{1,3}\\$?)(\\d+)");
        Matcher m = localPattern.matcher(temp);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String colPart = m.group(1);
            int rowNum = Integer.parseInt(m.group(2));
            int zeroBased = rowNum - 1;
            int newNum;
            if (rowMap.containsKey(zeroBased)) {
                newNum = rowMap.get(zeroBased) + 1;
            } else if (rowNum == oldLastDataRow + 1) {
                newNum = newDataLastRow + 1;
            } else if (rowNum > oldLastDataRow) {
                // 超出原数据区的引用（如 =SUM(H5:H999) 的万能上界）收紧到新文件末行，
                // 与桌面版 Python 实现保持完全一致
                newNum = newMaxRow;
            } else {
                newNum = rowNum;
            }
            m.appendReplacement(sb, Matcher.quoteReplacement(colPart + newNum));
        }
        m.appendTail(sb);
        temp = sb.toString();

        for (int i = 0; i < crossPh.size(); i++) {
            temp = temp.replace("__CROSS" + i + "__", crossPh.get(i));
        }
        return temp;
    }

    // ========================================================================
    //  工具方法
    // ========================================================================

    private Sheet resolveSheet(Workbook wb, String sheetName) {
        Sheet s = (sheetName != null && !sheetName.isEmpty()) ? wb.getSheet(sheetName) : null;
        return s != null ? s : (wb.getNumberOfSheets() > 0 ? wb.getSheetAt(0) : null);
    }

    /**
     * 动态识别数据起始行（0-based），兼容单行表头与多行表头两种格式。
     *
     * <p>判定：从表头区下一行起，第一个满足「非空单元格 >= 3」且「含数值单元格」的行。
     * 表头行即便有值也只是文本或公式，因此不会被命中。
     * 单行表头返回 2（第 3 行），三行表头返回 4（第 5 行）。
     */
    private int detectDataStartRow(Sheet sheet) {
        int maxRow = sheet.getLastRowNum();
        int maxCol = Math.max(computeMaxCol(sheet, maxRow), 1);
        for (int r = HEADER_ROW + 1; r <= maxRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            int nonEmpty = 0;
            boolean hasNumeric = false;
            for (int c = 0; c < maxCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null || cell.getCellType() == CellType.BLANK) continue;
                String txt = getCellStringValue(cell).trim();
                if (txt.isEmpty()) continue;
                nonEmpty++;
                if (cell.getCellType() == CellType.NUMERIC) hasNumeric = true;
            }
            if (nonEmpty >= 3 && hasNumeric) return r;
        }
        return HEADER_ROW + 1;
    }

    /** 表头区内最后一个非空列之后的位置（用作列循环上界） */
    private int headerAreaMaxCol(Sheet sheet, int dataStartRow) {
        int maxCol = 0;
        for (int r = HEADER_ROW; r < dataStartRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && !getCellStringValue(cell).trim().isEmpty()) {
                    maxCol = Math.max(maxCol, c + 1);
                }
            }
        }
        return maxCol;
    }

    /**
     * 构建「列名 -> 列号(0-based)」映射，逐列在表头区内取最后一个非空、非公式的值。
     *
     * <p>新版表中 A~F 的列名落在表头区首行（靠跨行合并撑着），G/H 落在末行；
     * 「总金额」「个数」等汇总标签与其公式都不参与列名判定。
     */
    private Map<String, Integer> buildHeaderMap(Sheet sheet, int dataStartRow) {
        Map<String, Integer> map = new LinkedHashMap<>();
        int maxCol = headerAreaMaxCol(sheet, dataStartRow);
        for (int c = 0; c < maxCol; c++) {
            String name = "";
            for (int r = HEADER_ROW; r < dataStartRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(c);
                if (cell == null) continue;
                String txt = getCellStringValue(cell).trim();
                if (txt.isEmpty() || txt.startsWith("=")) continue;
                name = txt;      // 持续覆盖，最终留下表头区里最靠下的文本
            }
            if (!name.isEmpty() && !map.containsKey(name)) map.put(name, c);
        }
        return map;
    }

    private int resolveSplitCol(Sheet sheet, String splitColName, int dataStartRow) {
        Map<String, Integer> headerMap = buildHeaderMap(sheet, dataStartRow);
        Integer idx = headerMap.get(splitColName == null ? "" : splitColName.trim());
        if (idx == null) throw new IllegalArgumentException("未找到列名：" + splitColName
                + "（已识别：" + headerMap.keySet() + "）");
        return idx;
    }

    private int resolveAmountCol(Sheet sheet, int dataStartRow) {
        Map<String, Integer> headerMap = buildHeaderMap(sheet, dataStartRow);
        for (String cand : AMOUNT_HEADER_CANDIDATES) {
            Integer idx = headerMap.get(cand);
            if (idx != null) return idx;
        }
        return -1;
    }

    private double sumAmount(Sheet sheet, List<Integer> rows, int amountCol) {
        if (amountCol < 0) return 0;
        double sum = 0;
        for (int r : rows) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(amountCol);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                sum += cell.getNumericCellValue();
            }
        }
        return sum;
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
            for (int mr = region.getFirstRow(); mr <= region.getLastRow(); mr++) {
                for (int mc = region.getFirstColumn(); mc <= region.getLastColumn(); mc++) {
                    map.put(mr + "_" + mc, val);
                }
            }
        }
        return map;
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

    /** 组合文件名：成员值按顺序用下划线连接，连续重复值只保留一次 */
    private String composeGroupName(List<Segment> segs) {
        List<String> uniq = new ArrayList<>();
        for (Segment s : segs) {
            if (uniq.isEmpty() || !uniq.get(uniq.size() - 1).equals(s.value)) uniq.add(s.value);
        }
        return String.join("_", uniq);
    }

    /**
     * 把若干段包含的<b>真实行号（0-based）</b>拼成任务参数，供 processSegmentFile 定位数据行。
     * 注意：不能用 Segment.index（那是段序号，仅用于前端组合拆分的引用）。
     */
    private String joinRowIds(List<Segment> segs) {
        List<String> ids = new ArrayList<>();
        for (Segment s : segs) {
            for (Integer r : s.rows) ids.add(String.valueOf(r));
        }
        return String.join(",", ids);
    }

    private List<Integer> parseRowIds(String joined) {
        List<Integer> ids = new ArrayList<>();
        if (joined == null || joined.isEmpty()) return ids;
        for (String part : joined.split(",")) {
            if (!part.trim().isEmpty()) ids.add(Integer.parseInt(part.trim()));
        }
        return ids;
    }

    /** 段索引 -> 段对象 */
    public List<Segment> resolveSegments(MultipartFile file, String sheetName, String colName) throws Exception {
        return getSegments(file, sheetName, colName);
    }

    /** 同名输出自动加序号：浙江锦途.xlsx -> 浙江锦途_2.xlsx */
    private String uniqueEntryName(String base, Set<String> used) {
        String name = base;
        int idx = 2;
        while (used.contains(name.toLowerCase())) {
            name = sanitizeFilename(base + "_" + idx);
            idx++;
        }
        used.add(name.toLowerCase());
        return name + ".xlsx";
    }

    private String getSafeSheetName(Workbook wb, String baseName) {
        String name = baseName;
        int idx = 1;
        while (wb.getSheet(name) != null) name = baseName + "_" + idx++;
        return name;
    }

    private String extractBaseName(String fileName) {
        if (fileName == null || fileName.isEmpty()) return "拆分结果";
        int dot = fileName.lastIndexOf('.');
        return dot > 0 ? fileName.substring(0, dot) : fileName;
    }

    /** 文件/条目名清洗（长度上限放宽，供 ZIP 条目名使用） */
    private String sanitizeFilename(String name) {
        if (name == null || name.isEmpty()) return "未命名";
        name = name.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').trim();
        for (char ch : "\\/:*?\"<>|".toCharArray()) name = name.replace(String.valueOf(ch), "_");
        while (name.endsWith(".") || name.endsWith(" ")) name = name.substring(0, name.length() - 1);
        if (name.isEmpty()) name = "未命名";
        if (name.length() > MAX_FILENAME_LEN) name = name.substring(0, MAX_FILENAME_LEN);
        return name;
    }

    /** Sheet 名清洗（Excel 上限 31 字符） */
    private String sanitizeSheetName(String name) {
        if (name == null || name.isEmpty()) return "Sheet1";
        for (char ch : "\\/:*?\"[]".toCharArray()) name = name.replace(String.valueOf(ch), "_");
        if (name.length() > MAX_SHEET_NAME_LEN) name = name.substring(0, MAX_SHEET_NAME_LEN);
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
