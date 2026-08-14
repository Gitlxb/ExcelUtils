package com.example.project.service.hr;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 对应 Python: rs_gzb_pipei.py（核心逻辑）+ rs_gzb_pipei_main.py（GUI）
 * 功能：工资表匹配（模式A：甲方匹配内部；模式B：内部互相匹配）
 *
 * 与 Python 最新版完全对齐的能力：
 * 1) 按关键列键值查找匹配（非行位置）
 * 2) 值 + 公式双维度对比（POI 单 Workbook 同时可取 getCellFormula() 与缓存值，等价于 Python 双加载）
 * 3) 自动发现工作表内多个数据段（关键列再次出现表头字段名即新段）
 * 4) 合计行检测 + 参与条件（两边都有合计且数据行数相等才参与对比）
 * 5) 报告未匹配人员、目标表多出人员、合计行参与情况
 */
@Service
public class HrGzbPipeiService {

    private static final double EPS = 1e-9;
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final String OUTPUT_NAME = "匹配结果.xlsx";

    /**
     * 获取 Excel 文件的所有工作表名称
     */
    public List<String> getSheetNames(MultipartFile file) throws Exception {
        File tempFile = saveToTemp(file);
        try (Workbook wb = WorkbookFactory.create(tempFile)) {
            List<String> names = new ArrayList<>();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                names.add(wb.getSheetName(i));
            }
            return names;
        } finally {
            tempFile.delete();
        }
    }

    /**
     * 预览指定工作表的前 N 行
     */
    public List<RowPreview> previewRows(MultipartFile file, String sheetName, int maxRows, int maxCols) throws Exception {
        File tempFile = saveToTemp(file);
        try (Workbook wb = WorkbookFactory.create(tempFile)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) throw new IllegalArgumentException("工作表不存在: " + sheetName);
            List<RowPreview> result = new ArrayList<>();
            int limit = Math.min(maxRows, sheet.getLastRowNum() + 1);
            for (int r = 0; r < limit; r++) {
                Row row = sheet.getRow(r);
                List<String> cells = new ArrayList<>();
                if (row != null) {
                    int colLimit = Math.min(maxCols, row.getLastCellNum());
                    for (int c = 0; c < colLimit; c++) {
                        cells.add(getCellStringValue(row.getCell(c)));
                    }
                }
                result.add(new RowPreview(r + 1, cells));
            }
            return result;
        } finally {
            tempFile.delete();
        }
    }

    /**
     * 读取指定行的表头字段
     */
    public List<String> getHeaders(MultipartFile file, String sheetName, int headerRow) throws Exception {
        File tempFile = saveToTemp(file);
        try (Workbook wb = WorkbookFactory.create(tempFile)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) throw new IllegalArgumentException("工作表不存在: " + sheetName);
            Row row = sheet.getRow(headerRow - 1);
            if (row == null) throw new IllegalArgumentException("表头行不存在: 第" + headerRow + "行");
            List<String> headers = new ArrayList<>();
            for (int c = 0; c < row.getLastCellNum(); c++) {
                String val = getCellStringValue(row.getCell(c)).trim();
                if (!val.isEmpty()) {
                    headers.add(val);
                }
            }
            return headers;
        } finally {
            tempFile.delete();
        }
    }

    /**
     * 模式A：甲方工资表匹配内部工资表
     */
    public ProcessResult processModeA(MultipartFile jiaFile, MultipartFile neiFile,
                                      String jiaSheet, String neiSheet,
                                      int jiaHeaderRow, int neiHeaderRow,
                                      String jiaKey, String neiKey,
                                      String fieldMapping, String outputDir) throws Exception {
        File jiaTemp = saveToTemp(jiaFile);
        File neiTemp = saveToTemp(neiFile);
        try (Workbook wbJia = WorkbookFactory.create(jiaTemp);
             Workbook wbNei = WorkbookFactory.create(neiTemp)) {
            Sheet wsJia = wbJia.getSheet(jiaSheet);
            Sheet wsNei = wbNei.getSheet(neiSheet);
            if (wsJia == null) throw new IllegalArgumentException("甲方工作表不存在: " + jiaSheet);
            if (wsNei == null) throw new IllegalArgumentException("内部工作表不存在: " + neiSheet);

            List<String[]> mapping = parseFieldMapping(fieldMapping);
            CompareResult cr = compareSheetsAndMark(
                    wbJia, wsJia, jiaHeaderRow,
                    wbNei, wsNei, neiHeaderRow,
                    jiaKey, neiKey, mapping);

            File outputFile = new File(outputDir, OUTPUT_NAME);
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                applyHighlight(wbNei, neiSheet, cr.highlight);
                wbNei.write(fos);
            }
            return new ProcessResult(outputFile.getAbsolutePath(), cr);
        } finally {
            jiaTemp.delete();
            neiTemp.delete();
        }
    }

    /**
     * 模式B：同一文件内两个工作表互相匹配
     */
    public ProcessResult processModeB(MultipartFile file,
                                      String sheetA, String sheetB,
                                      int headerA, int headerB,
                                      String keyA, String keyB,
                                      String fieldMapping, String outputDir) throws Exception {
        File tempFile = saveToTemp(file);
        try (Workbook wb = WorkbookFactory.create(tempFile)) {
            Sheet wsA = wb.getSheet(sheetA);
            Sheet wsB = wb.getSheet(sheetB);
            if (wsA == null) throw new IllegalArgumentException("工作表A不存在: " + sheetA);
            if (wsB == null) throw new IllegalArgumentException("工作表B不存在: " + sheetB);

            List<String[]> mapping = parseFieldMapping(fieldMapping);
            CompareResult cr = compareSheetsAndMark(
                    wb, wsA, headerA,
                    wb, wsB, headerB,
                    keyA, keyB, mapping);

            File outputFile = new File(outputDir, OUTPUT_NAME);
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                applyHighlight(wb, sheetB, cr.highlight);
                wb.write(fos);
            }
            return new ProcessResult(outputFile.getAbsolutePath(), cr);
        } finally {
            tempFile.delete();
        }
    }

    // ==================== 核心匹配算法（1:1 移植 Python compare_sheets_and_mark） ====================

    private CompareResult compareSheetsAndMark(Workbook wb1, Sheet ws1, int headerRow1,
                                               Workbook wb2, Sheet ws2, int headerRow2,
                                               String keyField1, String keyField2,
                                               List<String[]> mapping) {
        Map<String, Integer> headers1 = readHeaders(ws1, headerRow1);
        Map<String, Integer> headers2 = readHeaders(ws2, headerRow2);

        if (!headers1.containsKey(keyField1)) {
            throw new IllegalArgumentException("工作表「" + ws1.getSheetName() + "」中找不到关键列「" + keyField1 + "」");
        }
        if (!headers2.containsKey(keyField2)) {
            throw new IllegalArgumentException("工作表「" + ws2.getSheetName() + "」中找不到关键列「" + keyField2 + "」");
        }

        int keyCol1 = headers1.get(keyField1);
        int keyCol2 = headers2.get(keyField2);
        List<Integer> hcols1 = new ArrayList<>(headers1.values());
        List<Integer> hcols2 = new ArrayList<>(headers2.values());

        // 自动发现两个表的所有数据段
        List<Section> sections1 = findSections(ws1, headerRow1, keyCol1, keyField1, hcols1);
        List<Section> sections2 = findSections(ws2, headerRow2, keyCol2, keyField2, hcols2);

        // 构建目标表全段关键列索引：key -> 行号（1-based，首次出现映射）
        Map<String, Integer> targetIndex = new LinkedHashMap<>();
        Integer totalRowInTarget = null;
        for (Section s : sections2) {
            for (int row2 = s.header + 1; row2 <= s.last; row2++) {
                if (isDataTotalRow(ws2, row2, keyCol2, hcols2)) {
                    totalRowInTarget = row2;
                    continue;
                }
                String keyStr = readKeyString(ws2, row2, keyCol2);
                if (keyStr == null) continue;
                if (keyStr.isEmpty() || keyStr.equals("合计")) continue;
                targetIndex.putIfAbsent(keyStr, row2);
            }
        }

        // 统计基准表数据行数与是否有合计行
        boolean baseHasTotal = false;
        int baseDataRows = 0;
        for (Section s : sections1) {
            if (s.hasTotal) baseHasTotal = true;
            for (int row1 = s.header + 1; row1 <= s.last; row1++) {
                if (isDataTotalRow(ws1, row1, keyCol1, hcols1)) continue;
                String keyStr = readKeyString(ws1, row1, keyCol1);
                if (keyStr != null && !keyStr.isEmpty() && !keyStr.equals("合计")) {
                    baseDataRows++;
                }
            }
        }

        // 合计行参与条件：两边都有合计且数据行数相等
        boolean targetHasTotal = totalRowInTarget != null;
        int targetDataRows = targetIndex.size();
        boolean totalCanParticipate = baseHasTotal && targetHasTotal && (baseDataRows == targetDataRows);

        // 遍历基准表执行匹配与对比
        Set<CellCoord> highlight = new HashSet<>();
        int matchedCount = 0;
        List<String> unmatchedKeys = new ArrayList<>();
        Set<String> matchedTargetKeys = new HashSet<>();

        for (Section s : sections1) {
            for (int row1 = s.header + 1; row1 <= s.last; row1++) {
                boolean isTotalInBase = isDataTotalRow(ws1, row1, keyCol1, hcols1);
                String keyStr = readKeyString(ws1, row1, keyCol1);
                if (keyStr == null) keyStr = "";

                int row2;
                if (isTotalInBase) {
                    if (totalCanParticipate) {
                        row2 = totalRowInTarget;
                    } else {
                        continue;
                    }
                } else if (!keyStr.isEmpty() && !keyStr.equals("合计")) {
                    Integer r2 = targetIndex.get(keyStr);
                    if (r2 == null) {
                        unmatchedKeys.add(keyStr);
                        continue;
                    }
                    row2 = r2;
                    matchedTargetKeys.add(keyStr);
                } else {
                    continue;
                }

                matchedCount++;

                for (String[] pair : mapping) {
                    String f1 = pair[0];
                    String f2 = pair[1];
                    if (!headers1.containsKey(f1) || !headers2.containsKey(f2)) continue;
                    int col1 = headers1.get(f1);
                    int col2 = headers2.get(f2);
                    CellView cv1 = getCellView(getCell(ws1, row1, col1));
                    CellView cv2 = getCellView(getCell(ws2, row2, col2));
                    if (!compareCellFull(cv1, cv2)) {
                        highlight.add(new CellCoord(row2, col2));
                    }
                }
            }
        }

        List<String> extraTargetKeys = new ArrayList<>();
        for (String k : targetIndex.keySet()) {
            if (!matchedTargetKeys.contains(k)) extraTargetKeys.add(k);
        }

        TotalInfo totalInfo = new TotalInfo(
                totalCanParticipate, baseHasTotal, targetHasTotal,
                baseDataRows, targetDataRows, extraTargetKeys);

        return new CompareResult(highlight, matchedCount, unmatchedKeys, totalInfo);
    }

    // ==================== 工具方法 ====================

    private Cell getCell(Sheet sheet, int row1based, int col1based) {
        Row row = sheet.getRow(row1based - 1);
        if (row == null) return null;
        return row.getCell(col1based - 1);
    }

    /** 读取关键列单元格的字符串值（用于建索引与匹配） */
    private String readKeyString(Sheet sheet, int row1based, int keyCol) {
        Cell cell = getCell(sheet, row1based, keyCol);
        if (cell == null) return null;
        String v = getCellStringValue(cell).trim();
        return v.isEmpty() ? "" : v;
    }

    private Map<String, Integer> readHeaders(Sheet sheet, int headerRow) {
        Map<String, Integer> headers = new LinkedHashMap<>();
        Row row = sheet.getRow(headerRow - 1);
        if (row == null) return headers;
        for (int c = 0; c < row.getLastCellNum(); c++) {
            String val = getCellStringValue(row.getCell(c)).trim();
            if (!val.isEmpty()) {
                headers.put(val, c + 1); // 1-based
            }
        }
        return headers;
    }

    /** 从底部向上搜索有数据的最后一行（1-based） */
    private int findLastDataRow(Sheet sheet, int headerRow, Collection<Integer> headerCols) {
        int maxRow = sheet.getLastRowNum();
        Set<Integer> colsToCheck = new HashSet<>(headerCols);
        if (colsToCheck.isEmpty()) {
            Row hr = sheet.getRow(headerRow - 1);
            if (hr != null) {
                for (int c = 0; c < hr.getLastCellNum(); c++) colsToCheck.add(c + 1);
            }
        }
        for (int r = maxRow; r > headerRow; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int col : colsToCheck) {
                Cell cell = row.getCell(col - 1);
                if (cell != null) {
                    String val = getCellStringValue(cell).trim();
                    if (!val.isEmpty()) {
                        return r + 1; // 1-based
                    }
                }
            }
        }
        return headerRow;
    }

    /** 检测某行任意表头列是否含 =SUM() 公式 → 合计行 */
    private boolean isTotalRow(Sheet ws, int row1based, List<Integer> headerCols) {
        for (int col : headerCols) {
            Cell cell = getCell(ws, row1based, col);
            if (cell != null && cell.getCellType() == CellType.FORMULA) {
                String formula = cell.getCellFormula();
                if (formula != null && formula.trim().toUpperCase().startsWith("SUM(")) {
                    return true;
                }
            }
        }
        return false;
    }

    /** 合计行双条件：关键列为空或「合计」且行内有 =SUM() 公式 */
    private boolean isDataTotalRow(Sheet ws, int row1based, int keyCol, List<Integer> headerCols) {
        String keyStr = readKeyString(ws, row1based, keyCol);
        boolean keyBlank = (keyStr == null || keyStr.isEmpty() || keyStr.equals("合计"));
        return keyBlank && isTotalRow(ws, row1based, headerCols);
    }

    /** 自动发现工作表内的多个数据段（关键列再次出现表头字段名即新段） */
    private List<Section> findSections(Sheet ws, int headerRow, int keyCol,
                                       String keyFieldName, List<Integer> headerCols) {
        List<Section> sections = new ArrayList<>();
        int currentHeader = headerRow;
        int maxRow = ws.getLastRowNum() + 1; // 1-based

        int scan = headerRow + 1;
        while (scan <= maxRow) {
            String keyStr = readKeyString(ws, scan, keyCol);
            if (keyFieldName.equals(keyStr)) {
                int endRow = scan - 1;
                boolean hasTotal = false;
                if (endRow >= currentHeader + 1) {
                    hasTotal = isDataTotalRow(ws, endRow, keyCol, headerCols);
                }
                sections.add(new Section(currentHeader, endRow, hasTotal));
                currentHeader = scan;
            }
            scan++;
        }

        int lastRow = findLastDataRow(ws, currentHeader, headerCols);
        if (lastRow > currentHeader) {
            boolean hasTotal = isDataTotalRow(ws, lastRow, keyCol, headerCols);
            sections.add(new Section(currentHeader, lastRow, hasTotal));
        }

        if (sections.isEmpty()) {
            sections.add(new Section(headerRow, lastRow, false));
        }
        return sections;
    }

    /** 单元格视图：缓存值 + 公式字符串（公式维度对比用） */
    private static class CellView {
        Object value;
        String formula; // 非公式单元格为 null
    }

    private CellView getCellView(Cell cell) {
        CellView cv = new CellView();
        if (cell == null) return cv;

        if (cell.getCellType() == CellType.FORMULA) {
            cv.formula = cell.getCellFormula();
            try {
                switch (cell.getCachedFormulaResultType()) {
                    case NUMERIC:
                        cv.value = DateUtil.isCellDateFormatted(cell)
                                ? formatDate(cell) : cell.getNumericCellValue();
                        break;
                    case STRING:
                        cv.value = cell.getStringCellValue();
                        break;
                    case BOOLEAN:
                        cv.value = cell.getBooleanCellValue();
                        break;
                    default:
                        cv.value = null;
                }
            } catch (Exception e) {
                cv.value = null;
            }
            return cv;
        }

        switch (cell.getCellType()) {
            case STRING:
                String s = cell.getStringCellValue();
                cv.value = (s == null || s.trim().isEmpty()) ? null : s;
                break;
            case NUMERIC:
                cv.value = DateUtil.isCellDateFormatted(cell)
                        ? formatDate(cell) : cell.getNumericCellValue();
                break;
            case BOOLEAN:
                cv.value = cell.getBooleanCellValue();
                break;
            default:
                cv.value = null;
        }
        return cv;
    }

    /** 值 + 公式双维度比较，返回是否一致 */
    private boolean compareCellFull(CellView cv1, CellView cv2) {
        boolean valueOk = compareValues(cv1.value, cv2.value);
        boolean formulaOk = true;
        if (cv1.formula != null && cv2.formula != null) {
            formulaOk = cv1.formula.trim().equals(cv2.formula.trim());
        }
        return valueOk && formulaOk;
    }

    /** 数值容错 + 空值语义比较（对齐 Python _compare_values） */
    private boolean compareValues(Object v1, Object v2) {
        if (v1 == null && v2 == null) return true;
        // 左边空白，右边为 0 → 视为相同
        if (v1 == null) {
            Double d = tryParse(v2);
            return d != null && d == 0.0;
        }
        if (v2 == null) return false;

        Double n1 = tryParse(v1);
        Double n2 = tryParse(v2);
        if (n1 != null && n2 != null) {
            return Math.abs(n1 - n2) < EPS;
        }
        return String.valueOf(v1).trim().equals(String.valueOf(v2).trim());
    }

    private Double tryParse(Object v) {
        if (v instanceof Number) return ((Number) v).doubleValue();
        if (v == null) return null;
        try {
            return Double.parseDouble(v.toString().trim());
        } catch (NumberFormatException e) {
            return null;
        }
    }

    private String formatDate(Cell cell) {
        try {
            return cell.getLocalDateTimeCellValue().toLocalDate().format(DATE_FMT);
        } catch (Exception e) {
            try {
                return cell.getDateCellValue().toString();
            } catch (Exception e2) {
                return "";
            }
        }
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return formatDate(cell);
                }
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val) && !Double.isInfinite(val)) {
                    return String.valueOf((long) val);
                }
                return String.valueOf(val);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }

    private File saveToTemp(MultipartFile file) throws IOException {
        File temp = File.createTempFile("upload_", ".xlsx");
        file.transferTo(temp);
        return temp;
    }

    private List<String[]> parseFieldMapping(String json) throws Exception {
        if (json == null || json.trim().isEmpty()) return new ArrayList<>();
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readValue(json, new TypeReference<List<String[]>>() {});
    }

    /** 标红差异单元格：红字(FF0000) + 浅蓝底(CCE5FF)，跳过合并单元格非主格 */
    private void applyHighlight(Workbook wb, String sheetName, Set<CellCoord> cells) {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) return;

        for (CellCoord coord : cells) {
            int rIdx = coord.row - 1;
            int cIdx = coord.col - 1;
            if (isMergedCell(sheet, rIdx, cIdx)) continue;
            Row row = sheet.getRow(rIdx);
            if (row == null) continue;
            Cell cell = row.getCell(cIdx);
            if (cell == null) continue;

            CellStyle oldStyle = cell.getCellStyle();
            Font oldFont = wb.getFontAt(oldStyle.getFontIndexAsInt());
            Font newFont = wb.createFont();
            newFont.setFontName(oldFont.getFontName());
            newFont.setFontHeightInPoints(oldFont.getFontHeightInPoints());
            newFont.setBold(oldFont.getBold());
            newFont.setItalic(oldFont.getItalic());
            newFont.setUnderline(oldFont.getUnderline());
            newFont.setStrikeout(oldFont.getStrikeout());
            newFont.setColor(IndexedColors.RED.getIndex());

            CellStyle newStyle = wb.createCellStyle();
            newStyle.cloneStyleFrom(oldStyle);
            newStyle.setFont(newFont);
            if (wb instanceof XSSFWorkbook) {
                ((XSSFCellStyle) newStyle).setFillForegroundColor(
                        new XSSFColor(new byte[]{(byte) 0xCC, (byte) 0xE5, (byte) 0xFF}, null));
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else {
                newStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            cell.setCellStyle(newStyle);
        }
    }

    private boolean isMergedCell(Sheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(row, col)) {
                return region.getFirstRow() != row || region.getFirstColumn() != col;
            }
        }
        return false;
    }

    // ==================== DTO ====================

    public static class RowPreview {
        public int rowNum;
        public List<String> cells;
        public RowPreview(int rowNum, List<String> cells) {
            this.rowNum = rowNum;
            this.cells = cells;
        }
    }

    /** 单段（section）：表头行 / 数据尾行 / 是否含合计行 */
    private static class Section {
        int header;
        int last;
        boolean hasTotal;
        Section(int header, int last, boolean hasTotal) {
            this.header = header;
            this.last = last;
            this.hasTotal = hasTotal;
        }
    }

    /** 合计行参与情况汇总 */
    public static class TotalInfo {
        public boolean participated;
        public boolean baseHasTotal;
        public boolean targetHasTotal;
        public int baseDataRows;
        public int targetDataRows;
        public List<String> extraTargetKeys;
        public TotalInfo(boolean participated, boolean baseHasTotal, boolean targetHasTotal,
                         int baseDataRows, int targetDataRows, List<String> extraTargetKeys) {
            this.participated = participated;
            this.baseHasTotal = baseHasTotal;
            this.targetHasTotal = targetHasTotal;
            this.baseDataRows = baseDataRows;
            this.targetDataRows = targetDataRows;
            this.extraTargetKeys = extraTargetKeys;
        }
    }

    /** 匹配结果 */
    public static class CompareResult {
        public final Set<CellCoord> highlight;
        public final int matchedCount;
        public final List<String> unmatchedKeys;
        public final TotalInfo totalInfo;
        public CompareResult(Set<CellCoord> highlight, int matchedCount,
                             List<String> unmatchedKeys, TotalInfo totalInfo) {
            this.highlight = highlight;
            this.matchedCount = matchedCount;
            this.unmatchedKeys = unmatchedKeys;
            this.totalInfo = totalInfo;
        }
    }

    public static class ProcessResult {
        public final String outputPath;
        public final int matchedCount;
        public final int diffCount;
        public final List<String> unmatchedKeys;
        public final TotalInfo totalInfo;
        public ProcessResult(String outputPath, CompareResult cr) {
            this.outputPath = outputPath;
            this.matchedCount = cr.matchedCount;
            this.diffCount = cr.highlight.size();
            this.unmatchedKeys = cr.unmatchedKeys;
            this.totalInfo = cr.totalInfo;
        }
    }

    private static class CellCoord {
        final int row;
        final int col;
        CellCoord(int row, int col) { this.row = row; this.col = col; }
        @Override public int hashCode() { return Objects.hash(row, col); }
        @Override public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof CellCoord)) return false;
            CellCoord c = (CellCoord) o;
            return row == c.row && col == c.col;
        }
    }
}
