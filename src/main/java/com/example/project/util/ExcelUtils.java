package com.example.project.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import java.util.HashSet;
import java.util.Set;

/**
 * Excel 公共工具类
 * 提供 Excel 列字母转换、工作表名称处理等通用工具方法
 */
public class ExcelUtils {

    /** Excel 非法字符（不能出现在工作表名称中） */
    private static final Set<Character> FORBIDDEN_CHARS = new HashSet<>();
    static {
        for (char c : new char[]{'\\', '/', '*', '[', ']', ':', '?'}) {
            FORBIDDEN_CHARS.add(c);
        }
    }

    /**
     * 将 Excel 列字母转换为 0-based 列索引
     * A→0, B→1, Z→25, AA→26, AB→27 ...
     */
    public static int excelColumnToIndex(String columnLetter) {
        if (columnLetter == null || columnLetter.isEmpty()) {
            throw new IllegalArgumentException("列字母不能为空");
        }
        int index = 0;
        String upper = columnLetter.toUpperCase();
        for (int i = 0; i < upper.length(); i++) {
            char c = upper.charAt(i);
            if (c < 'A' || c > 'Z') {
                throw new IllegalArgumentException("无效的列字母: " + columnLetter);
            }
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }

    /**
     * 将 0-based 列索引转换为 Excel 列字母
     * 0→A, 1→B, 25→Z, 26→AA ...
     */
    public static String indexToExcelColumn(int index) {
        if (index < 0) {
            throw new IllegalArgumentException("列索引不能为负数: " + index);
        }
        StringBuilder sb = new StringBuilder();
        int temp = index;
        while (temp >= 0) {
            sb.insert(0, (char) ((temp % 26) + 'A'));
            temp = (temp / 26) - 1;
        }
        return sb.toString();
    }

    /**
     * 清理工作表名称：替换非法字符、截断长度至31字符
     */
    public static String cleanSheetName(String name) {
        if (name == null || name.isEmpty()) {
            return "空值";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < name.length(); i++) {
            char c = name.charAt(i);
            if (FORBIDDEN_CHARS.contains(c)) {
                sb.append('_');
            } else {
                sb.append(c);
            }
        }
        if (sb.length() > 31) {
            sb.setLength(31);
        }
        return sb.toString();
    }

    /**
     * 判断工作表名称是否合法
     */
    public static boolean isValidSheetName(String name) {
        if (name == null || name.isEmpty() || name.length() > 31) {
            return false;
        }
        for (int i = 0; i < name.length(); i++) {
            if (FORBIDDEN_CHARS.contains(name.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    /**
     * 确保工作表名称唯一（如果重复则添加数字后缀）
     */
    public static String getUniqueSheetName(String baseName, Set<String> existingNames) {
        String name = cleanSheetName(baseName);
        String original = name;
        int counter = 1;
        while (!isValidSheetName(name) || existingNames.contains(name)) {
            String suffix = "_" + counter;
            int maxLen = 31 - suffix.length();
            if (maxLen <= 0) {
                name = suffix;
            } else {
                name = (original.length() > maxLen ? original.substring(0, maxLen) : original) + suffix;
            }
            counter++;
        }
        return name;
    }

    /**
     * 获取单元格的字符串值（安全取值，不会抛异常）
     */
    public static String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case ERROR:
                return "";
            case FORMULA:
                try {
                    CellType cachedType = cell.getCachedFormulaResultType();
                    if (cachedType == CellType.STRING) {
                        return cell.getStringCellValue();
                    } else if (cachedType == CellType.NUMERIC) {
                        return String.valueOf(cell.getNumericCellValue());
                    } else if (cachedType == CellType.BOOLEAN) {
                        return String.valueOf(cell.getBooleanCellValue());
                    } else if (cachedType == CellType.ERROR) {
                        return "";
                    }
                    return "";
                } catch (Exception e) {
                    return "";
                }
            default:
                return "";
        }
    }

    /**
     * 获取单元格的数字值（安全取值，不会抛异常）
     */
    public static double getCellDoubleValue(Cell cell) {
        if (cell == null) return 0;
        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    String strValue = cell.getStringCellValue();
                    if (strValue != null && !strValue.isEmpty()) {
                        return Double.parseDouble(strValue);
                    }
                    return 0;
                case ERROR:
                    return 0;
                case FORMULA:
                    CellType cachedType = cell.getCachedFormulaResultType();
                    if (cachedType == CellType.NUMERIC) {
                        return cell.getNumericCellValue();
                    } else if (cachedType == CellType.STRING) {
                        String s = cell.getStringCellValue();
                        if (s != null && !s.isEmpty()) {
                            return Double.parseDouble(s);
                        }
                    }
                    return 0;
                default:
                    return 0;
            }
        } catch (Exception e) {
            return 0;
        }
    }

    /**
     * 复制单元格值（源→目标）
     */
    public static void copyCellValue(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) return;
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                targetCell.setCellType(CellType.BLANK);
                break;
        }
    }

    /**
     * 复制行数据（源→目标行，指定列范围）
     */
    public static void copyRow(Row sourceRow, Row targetRow, int startCol, int endCol) {
        if (sourceRow == null || targetRow == null) return;
        for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
            Cell sourceCell = sourceRow.getCell(colIndex);
            Cell targetCell = targetRow.createCell(colIndex);
            if (sourceCell != null) {
                copyCellValue(sourceCell, targetCell);
            }
        }
    }

    /**
     * 创建表格头部样式（灰色背景 + 加粗）
     */
    public static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        ((XSSFCellStyle)style).setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x44, (byte)0x72, (byte)0xC4}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 创建黄色填充样式
     */
    public static CellStyle createYellowFillStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    /**
     * 创建红色填充样式（用于标记重复值）
     */
    public static CellStyle createDuplicateFillStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
