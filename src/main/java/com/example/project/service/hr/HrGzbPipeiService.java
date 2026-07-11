package com.example.project.service.hr;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

/**
 * 对应 Python: rs_gzb_pipei.py + rs_gzb_pipei_main.py
 * 功能：工资表匹配（模式A：甲方匹配内部；模式B：内部互相匹配）
 */
@Service
public class HrGzbPipeiService {

    private static final double EPS = 1e-9;

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
     * @param jiaFile       甲方文件
     * @param neiFile       内部文件
     * @param jiaSheet      甲方工作表
     * @param neiSheet      内部工作表
     * @param jiaHeaderRow  甲方表头行（1-based）
     * @param neiHeaderRow  内部表头行（1-based）
     * @param fieldMapping  JSON 格式的字段映射 [[甲方字段, 内部字段], ...]
     * @param outputDir     输出目录
     * @return 输出文件路径
     */
    public ProcessResult processModeA(MultipartFile jiaFile, MultipartFile neiFile,
                               String jiaSheet, String neiSheet,
                               int jiaHeaderRow, int neiHeaderRow,
                               String fieldMapping, String outputDir) throws Exception {
        File jiaTemp = saveToTemp(jiaFile);
        File neiTemp = saveToTemp(neiFile);
        try {
            List<String[]> mapping = parseFieldMapping(fieldMapping);

            // 1. data_only=true 对比（获取计算值）
            Set<CellCoord> highlightCells;
            int compareRows;
            try (Workbook wbJia = WorkbookFactory.create(jiaTemp);
                 Workbook wbNei = WorkbookFactory.create(neiTemp)) {
                Sheet wsJia = wbJia.getSheet(jiaSheet);
                Sheet wsNei = wbNei.getSheet(neiSheet);
                if (wsJia == null) throw new IllegalArgumentException("甲方工作表不存在");
                if (wsNei == null) throw new IllegalArgumentException("内部工作表不存在");

                Map<String, Integer> headers1 = readHeaders(wsJia, jiaHeaderRow);
                Map<String, Integer> headers2 = readHeaders(wsNei, neiHeaderRow);

                int lastRow1 = findLastDataRow(wsJia, jiaHeaderRow, headers1.values());
                int lastRow2 = findLastDataRow(wsNei, neiHeaderRow, headers2.values());
                int dataRows1 = lastRow1 - jiaHeaderRow;
                int dataRows2 = lastRow2 - neiHeaderRow;
                compareRows = Math.min(dataRows1, dataRows2);

                highlightCells = new HashSet<>();
                for (int i = 0; i < compareRows; i++) {
                    int row1 = jiaHeaderRow + 1 + i;
                    int row2 = neiHeaderRow + 1 + i;
                    for (String[] pair : mapping) {
                        String f1 = pair[0];
                        String f2 = pair[1];
                        Integer col1 = headers1.get(f1);
                        Integer col2 = headers2.get(f2);
                        if (col1 == null || col2 == null) continue;
                        Object v1 = getCellValueForCompare(wsJia.getRow(row1 - 1).getCell(col1 - 1));
                        Object v2 = getCellValueForCompare(wsNei.getRow(row2 - 1).getCell(col2 - 1));
                        if (!compareValues(v1, v2)) {
                            highlightCells.add(new CellCoord(row2, col2));
                        }
                    }
                }
            }

            // 2. 普通模式加载内部文件（保留公式和格式）
            File outputFile = new File(outputDir, "匹配结果_" + System.currentTimeMillis() + ".xlsx");
            try (Workbook wbOut = WorkbookFactory.create(neiTemp);
                 FileOutputStream fos = new FileOutputStream(outputFile)) {
                applyHighlight(wbOut, neiSheet, highlightCells);
                wbOut.write(fos);
            }

            return new ProcessResult(outputFile.getAbsolutePath(), compareRows, highlightCells.size());
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
                               String fieldMapping, String outputDir) throws Exception {
        File tempFile = saveToTemp(file);
        try {
            List<String[]> mapping = parseFieldMapping(fieldMapping);

            Set<CellCoord> highlightCells;
            int compareRows;
            try (Workbook wbData = WorkbookFactory.create(tempFile)) {
                Sheet wsA = wbData.getSheet(sheetA);
                Sheet wsB = wbData.getSheet(sheetB);
                if (wsA == null) throw new IllegalArgumentException("工作表A不存在");
                if (wsB == null) throw new IllegalArgumentException("工作表B不存在");

                Map<String, Integer> headersA = readHeaders(wsA, headerA);
                Map<String, Integer> headersB = readHeaders(wsB, headerB);

                int lastRowA = findLastDataRow(wsA, headerA, headersA.values());
                int lastRowB = findLastDataRow(wsB, headerB, headersB.values());
                int dataRowsA = lastRowA - headerA;
                int dataRowsB = lastRowB - headerB;
                compareRows = Math.min(dataRowsA, dataRowsB);

                highlightCells = new HashSet<>();
                for (int i = 0; i < compareRows; i++) {
                    int rowA = headerA + 1 + i;
                    int rowB = headerB + 1 + i;
                    for (String[] pair : mapping) {
                        String fA = pair[0];
                        String fB = pair[1];
                        Integer colA = headersA.get(fA);
                        Integer colB = headersB.get(fB);
                        if (colA == null || colB == null) continue;
                        Object vA = getCellValueForCompare(wsA.getRow(rowA - 1).getCell(colA - 1));
                        Object vB = getCellValueForCompare(wsB.getRow(rowB - 1).getCell(colB - 1));
                        if (!compareValues(vA, vB)) {
                            highlightCells.add(new CellCoord(rowB, colB));
                        }
                    }
                }
            }

            File outputFile = new File(outputDir, "匹配结果_" + System.currentTimeMillis() + ".xlsx");
            try (Workbook wbOut = WorkbookFactory.create(tempFile);
                 FileOutputStream fos = new FileOutputStream(outputFile)) {
                applyHighlight(wbOut, sheetB, highlightCells);
                wbOut.write(fos);
            }

            return new ProcessResult(outputFile.getAbsolutePath(), compareRows, highlightCells.size());
        } finally {
            tempFile.delete();
        }
    }

    // ==================== 内部工具方法 ====================

    private File saveToTemp(MultipartFile file) throws IOException {
        File temp = File.createTempFile("upload_", ".xlsx");
        file.transferTo(temp);
        return temp;
    }

    private List<String[]> parseFieldMapping(String json) throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readValue(json, new TypeReference<List<String[]>>() {});
    }

    private Map<String, Integer> readHeaders(Sheet sheet, int headerRow) {
        Map<String, Integer> headers = new HashMap<>();
        Row row = sheet.getRow(headerRow - 1);
        if (row == null) return headers;
        for (int c = 0; c < row.getLastCellNum(); c++) {
            String val = getCellStringValue(row.getCell(c)).trim();
            if (!val.isEmpty()) {
                headers.put(val, c + 1); // 1-based column index
            }
        }
        return headers;
    }

    private int findLastDataRow(Sheet sheet, int headerRow, Collection<Integer> headerCols) {
        int maxRow = sheet.getLastRowNum();
        Set<Integer> colsToCheck = new HashSet<>(headerCols);
        if (colsToCheck.isEmpty()) {
            for (int c = 0; c < sheet.getRow(headerRow - 1).getLastCellNum(); c++) {
                colsToCheck.add(c + 1);
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

    private Object getCellValueForCompare(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                String s = cell.getStringCellValue().trim();
                return s.isEmpty() ? null : s;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    return cell.getNumericCellValue();
                } catch (Exception e) {
                    try {
                        String fs = cell.getStringCellValue().trim();
                        return fs.isEmpty() ? null : fs;
                    } catch (Exception e2) {
                        return null;
                    }
                }
            default:
                return null;
        }
    }

    private boolean compareValues(Object v1, Object v2) {
        if (v1 == null && v2 == null) return true;
        // 左边空白，右边为0 → 视为相同
        if (v1 == null && v2 instanceof Number) {
            return ((Number) v2).doubleValue() == 0;
        }
        if (v1 == null || v2 == null) return false;

        // 数值比较
        if (v1 instanceof Number && v2 instanceof Number) {
            return Math.abs(((Number) v1).doubleValue() - ((Number) v2).doubleValue()) < EPS;
        }
        try {
            double n1 = Double.parseDouble(v1.toString());
            double n2 = Double.parseDouble(v2.toString());
            return Math.abs(n1 - n2) < EPS;
        } catch (NumberFormatException ignored) {}

        return v1.toString().trim().equals(v2.toString().trim());
    }

    private void applyHighlight(Workbook wb, String sheetName, Set<CellCoord> cells) {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) return;

        // 创建标红样式（红色字体 + 浅蓝色背景）
        CellStyle highlightStyle = wb.createCellStyle();
        Font redFont = wb.createFont();
        redFont.setColor(IndexedColors.RED.getIndex());
        highlightStyle.setFont(redFont);
        if (wb instanceof XSSFWorkbook) {
            ((XSSFCellStyle) highlightStyle).setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xCC, (byte) 0xE5, (byte) 0xFF}, null));
            highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } else {
            highlightStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        for (CellCoord coord : cells) {
            int rIdx = coord.row - 1;
            int cIdx = coord.col - 1;
            // 跳过合并单元格的非主单元格（与Python MergedCell处理对齐）
            if (isMergedCell(sheet, rIdx, cIdx)) continue;
            Row row = sheet.getRow(rIdx);
            if (row == null) continue;
            Cell cell = row.getCell(cIdx);
            if (cell == null) continue;

            // 保留原有字体属性，只改颜色
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
                ((XSSFCellStyle) newStyle).setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xCC, (byte) 0xE5, (byte) 0xFF}, null));
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

    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val)) {
                    return String.valueOf((long) val);
                }
                return String.valueOf(val);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try { return cell.getStringCellValue(); }
                catch (Exception e) {
                    try { return String.valueOf(cell.getNumericCellValue()); }
                    catch (Exception e2) { return ""; }
                }
            default: return "";
        }
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

    public static class ProcessResult {
        public final String outputPath;
        public final int compareRows;
        public final int diffCount;
        public ProcessResult(String outputPath, int compareRows, int diffCount) {
            this.outputPath = outputPath;
            this.compareRows = compareRows;
            this.diffCount = diffCount;
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
