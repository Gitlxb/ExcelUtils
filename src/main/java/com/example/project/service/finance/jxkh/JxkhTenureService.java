package com.example.project.service.finance.jxkh;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 绩效考核 - 员工在职月份计算（tenure_months_calculator.py 转 Java）
 *
 * 业务规则（与 Python 原版一致）：
 * 1. 精确识别 5 个固定表头：状态 / 姓名 / 身份证号码 / 入职日期 / 离职日期。
 * 2. 分组键 = 姓名 + 身份证号码（同名不同身份证视为不同人）。
 * 3. 行类型：在职（忽略离职日期）、离职（有入职+离职）、其他（原样保留）。
 * 4. 主在职区间：起点 = max(在职最早入职, 离职最早入职) + 起算偏移（默认从次日算起）；
 *    终点 = min(统计月末, 在职截止日)。
 * 5. 短离职区间：每条离职行 [入职起算日, 离职日]，从主在职区间减去其并集。
 * 6. 按 (年,月) 聚合剩余天数，展开到在职行；离职行按自身短离职区间展开。
 * 7. 新增列：在职月份（日期单元格 yyyy/m）/ 在职天数 / 总在职天数（按人合并）。
 */
@Service
public class JxkhTenureService {

    private static final DateTimeFormatter MONTH_FMT = DateTimeFormatter.ofPattern("yyyy-MM");
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter[] DATE_PARSERS = {
            DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("yyyy/MM"),
            DateTimeFormatter.ofPattern("yyyy-MM"),
            DateTimeFormatter.ofPattern("yyyy年MM月dd日"),
            DateTimeFormatter.ofPattern("yyyyMMdd"),
            DateTimeFormatter.ofPattern("yyyy.MM.dd")
    };

    // 5 个固定表头（精确匹配）
    private static final String COL_STATUS = "状态";
    private static final String COL_EMP = "姓名";
    private static final String COL_ID = "身份证号码";
    private static final String COL_HIRE = "入职日期";
    private static final String COL_DEPART = "离职日期";

    /** 计算结果：输出文件路径 + 处理日志 */
    public static class AnalyzeResult {
        public String filePath;
        public List<String> logs = new ArrayList<>();
    }

    // ============================ 入口 ============================
    public AnalyzeResult analyze(MultipartFile file, String sheetName, String statMonth,
                                 String asOfDate, boolean mergeTotalDays, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        try {
            return doAnalyze(tempFile, sheetName, statMonth, asOfDate, mergeTotalDays, uploadDir);
        } finally {
            tempFile.delete();
        }
    }

    private AnalyzeResult doAnalyze(File inputFile, String sheetName, String statMonth,
                                    String asOfDate, boolean mergeTotalDays, String uploadDir) throws Exception {
        List<String> logs = new ArrayList<>();
        AnalyzeResult result = new AnalyzeResult();
        result.logs = logs;

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook wb = WorkbookFactory.create(fis)) {

            Sheet sheet = (sheetName != null && !sheetName.isEmpty())
                    ? wb.getSheet(sheetName) : wb.getSheetAt(0);
            if (sheet == null) throw new Exception("工作表不存在：" + sheetName);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new Exception("没有表头行");

            // ---- 精确识别表头 ----
            Map<String, Integer> colMap = new LinkedHashMap<>();
            List<String> origCols = new ArrayList<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                String h = ExcelUtils.getCellStringValue(headerRow.getCell(c)).trim();
                if (!h.isEmpty()) {
                    colMap.put(h, c);
                    origCols.add(h);
                }
            }

            Integer idxStatus = colMap.get(COL_STATUS);
            Integer idxEmp = colMap.get(COL_EMP);
            Integer idxId = colMap.get(COL_ID);
            Integer idxHire = colMap.get(COL_HIRE);
            Integer idxDepart = colMap.get(COL_DEPART);

            logs.add("表头识别结果：状态=" + idxStatus + ", 姓名=" + idxEmp
                    + ", 身份证=" + idxId + ", 入职=" + idxHire + ", 离职=" + idxDepart);
            logs.add("最终使用表头：状态=" + COL_STATUS + ", 姓名=" + COL_EMP
                    + ", 身份证=" + COL_ID + ", 入职=" + COL_HIRE + ", 离职=" + COL_DEPART);

            if (idxStatus == null) {
                logs.add("❌ 未识别到『状态』列，请检查表头是否为『状态』。");
                throw new Exception("未识别到『状态』列");
            }
            if (idxHire == null) {
                logs.add("❌ 未识别到『入职日期』列，请检查表头是否为『入职日期』。");
                throw new Exception("未识别到『入职日期』列");
            }
            if (idxEmp == null) logs.add("⚠️ 未识别到『姓名』列。");
            if (idxId == null) logs.add("⚠️ 未识别到『身份证号码』列，将仅按姓名分组。");

            // ---- 解析统计月末与在职截止日 ----
            YearMonth statYm = YearMonth.parse(statMonth.trim(), MONTH_FMT);
            LocalDate statFirst = statYm.atDay(1);
            LocalDate statEnd = statYm.atEndOfMonth();

            LocalDate asOf;
            if (asOfDate == null || asOfDate.trim().isEmpty() || "today".equalsIgnoreCase(asOfDate.trim())) {
                asOf = LocalDate.now();
            } else {
                asOf = parseDate(asOfDate.trim());
                if (asOf == null) throw new Exception("在职截止日格式不正确：" + asOfDate);
            }

            // 含入职日：固定默认 false（从入职次日算起），界面不暴露该参数
            boolean includeHireDay = false;
            logs.add("统计截止月份：" + statMonth + " | 含入职日：" + includeHireDay);
            logs.add("在职天数截止日(as_of)：" + asOf + "（主在职终点不超过该日）");

            // ---- 读取数据行 ----
            List<RawRow> rows = new ArrayList<>();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                RawRow rr = new RawRow();
                for (String col : origCols) {
                    rr.cells.put(col, readCellValue(row.getCell(colMap.get(col)), col));
                }
                rr.status = strVal(rr.cells.get(COL_STATUS)).trim();
                rr.name = strVal(rr.cells.get(COL_EMP));
                rr.id = strVal(rr.cells.get(COL_ID));
                rr.hire = parseDateValue(rr.cells.get(COL_HIRE));
                rr.depart = (idxDepart != null) ? parseDateValue(rr.cells.get(COL_DEPART)) : null;
                rows.add(rr);
            }
            logs.add("原始数据行数：" + rows.size());

            // ---- 按 (姓名 + 身份证) 分组 ----
            Map<String, List<RawRow>> groups = new LinkedHashMap<>();
            for (RawRow rr : rows) {
                groups.computeIfAbsent(personKey(rr.name, rr.id), k -> new ArrayList<>()).add(rr);
            }

            // ---- 逐人计算并展开 ----
            List<OutRow> outRows = new ArrayList<>();
            int nEmployed = 0, nLeft = 0, nOther = 0, nMissingHire = 0, nNoMonths = 0;

            for (List<RawRow> members : groups.values()) {
                List<LocalDate> employedHires = new ArrayList<>();
                List<Interval> leaveIntervals = new ArrayList<>();
                List<RawRow> employedRows = new ArrayList<>();
                List<RawRow> otherRows = new ArrayList<>();
                List<LeaveRec> leaveRecs = new ArrayList<>();

                for (RawRow rr : members) {
                    if ("在职".equals(rr.status)) {
                        nEmployed++;
                        employedRows.add(rr);
                        if (rr.hire != null) employedHires.add(rr.hire);
                        else nMissingHire++;
                    } else if ("离职".equals(rr.status)) {
                        nLeft++;
                        if (rr.hire == null) {
                            leaveRecs.add(new LeaveRec(rr, null, null));
                        } else {
                            LocalDate leaveEnd = (rr.depart != null) ? minDate(rr.depart, statEnd) : statEnd;
                            Interval iv = computeLeaveInterval(rr.hire, leaveEnd, includeHireDay);
                            leaveIntervals.add(iv);
                            leaveRecs.add(new LeaveRec(rr, iv.start, iv.end));
                        }
                    } else {
                        nOther++;
                        otherRows.add(rr);
                    }
                }

                Map<YM, Integer> mainDays = computeMainEmployment(
                        employedHires, leaveIntervals, statEnd, asOf, includeHireDay);

                // 在职行
                if (!employedRows.isEmpty()) {
                    if (mainDays.isEmpty()) nNoMonths++;
                    if (mainDays.isEmpty()) {
                        for (RawRow rr : employedRows) {
                            OutRow or = new OutRow(rr);
                            or.tenureMonth = null;
                            or.tenureDays = 0;
                            outRows.add(or);
                        }
                    } else {
                        List<YM> sorted = new ArrayList<>(mainDays.keySet());
                        sorted.sort(Comparator.comparingLong(YM::key));
                        for (YM ym : sorted) {
                            for (RawRow rr : employedRows) {
                                OutRow or = new OutRow(rr);
                                or.tenureMonth = LocalDate.of(ym.year, ym.month, 1);
                                or.tenureDays = mainDays.get(ym);
                                outRows.add(or);
                            }
                        }
                    }
                }

                // 离职行
                for (LeaveRec lr : leaveRecs) {
                    if (lr.start == null || lr.start.isAfter(lr.end)) {
                        OutRow or = new OutRow(lr.row);
                        if (lr.start == null) {
                            or.tenureMonth = null;
                        } else {
                            LocalDate h = lr.row.hire;
                            or.tenureMonth = (h != null) ? LocalDate.of(h.getYear(), h.getMonthValue(), 1) : null;
                        }
                        or.tenureDays = 0;
                        outRows.add(or);
                    } else {
                        Map<YM, Integer> md = daysByMonth(lr.start, lr.end);
                        List<YM> sorted = new ArrayList<>(md.keySet());
                        sorted.sort(Comparator.comparingLong(YM::key));
                        for (YM ym : sorted) {
                            OutRow or = new OutRow(lr.row);
                            or.tenureMonth = LocalDate.of(ym.year, ym.month, 1);
                            or.tenureDays = md.get(ym);
                            outRows.add(or);
                        }
                    }
                }

                // 其他状态行
                for (RawRow rr : otherRows) {
                    OutRow or = new OutRow(rr);
                    or.tenureMonth = null;
                    or.tenureDays = 0;
                    outRows.add(or);
                }
            }

            // ---- 汇总日志 ----
            long validRows = outRows.stream().filter(o -> o.tenureMonth != null).count();
            logs.add("✅ 计算完成");
            logs.add("   - 在职（参与主在职计算）：" + nEmployed + " 行");
            logs.add("   - 离职（按行展开短离职）：" + nLeft + " 行");
            logs.add("   - 其他状态（原样保留）：" + nOther + " 行");
            logs.add("   展开后总明细：" + outRows.size() + " 行");
            logs.add("   有在职月份的有效明细：" + validRows + " 行");
            logs.add("   - 在职行中入职日期缺失：" + nMissingHire + " 行");
            logs.add("   - 在职行中无在职月份：" + nNoMonths + " 行");
            logs.add("   工作表『在职月份结果』列：原表列 + 在职月份 + 在职天数 + 总在职天数");
            logs.add(mergeTotalDays ? "   已按人合并：仅『总在职天数』列（垂直居中）"
                                    : "   未合并：总在职天数列按行展开（每行 = 该行在职天数）");

            // ---- 写出 Excel ----
            File outFile = new File(uploadDir, "员工在职月份结果.xlsx");
            writeOutput(origCols, outRows, mergeTotalDays, idxEmp, idxId, outFile, logs);
            result.filePath = outFile.getAbsolutePath();
        }
        return result;
    }

    // ============================ 写出（单工作表） ============================
    private void writeOutput(List<String> origCols, List<OutRow> outRows,
                             boolean mergeTotalDays, Integer idxEmp, Integer idxId,
                             File outFile, List<String> logs) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("在职月份结果");

            List<String> outCols = new ArrayList<>(origCols);
            outCols.add("在职月份");
            outCols.add("在职天数");
            outCols.add("总在职天数");

            // 表头（清除加粗/边框，与 Python 一致）
            Row header = sheet.createRow(0);
            CellStyle plainHeader = wb.createCellStyle();
            for (int c = 0; c < outCols.size(); c++) {
                Cell cell = header.createCell(c);
                cell.setCellValue(outCols.get(c));
                cell.setCellStyle(plainHeader);
            }

            // 共享单元格样式（避免逐格新建导致 Excel 样式数超限）
            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-MM-dd"));
            CellStyle monthStyle = wb.createCellStyle();
            monthStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy/m"));

            // 计算「总在职天数」
            Map<OutRow, Integer> totalDaysMap = new LinkedHashMap<>();
            if (mergeTotalDays && idxEmp != null && idxId != null) {
                Map<String, Integer> grpSum = new LinkedHashMap<>();
                for (OutRow or : outRows) {
                    String key = personKey(strVal(or.cells.get(COL_EMP)), strVal(or.cells.get(COL_ID)));
                    grpSum.merge(key, or.tenureDays, Integer::sum);
                }
                for (OutRow or : outRows) {
                    String key = personKey(strVal(or.cells.get(COL_EMP)), strVal(or.cells.get(COL_ID)));
                    totalDaysMap.put(or, grpSum.get(key));
                }
            } else {
                for (OutRow or : outRows) totalDaysMap.put(or, or.tenureDays);
            }

            int r = 1;
            for (OutRow or : outRows) {
                Row row = sheet.createRow(r++);
                for (int c = 0; c < origCols.size(); c++) {
                    String col = origCols.get(c);
                    Cell cell = row.createCell(c);
                    writeOriginalCell(cell, or.cells.get(col), dateStyle);
                }
                int baseC = origCols.size();
                // 在职月份：日期单元格 yyyy/m
                Cell mCell = row.createCell(baseC);
                if (or.tenureMonth != null) {
                    mCell.setCellValue(or.tenureMonth);
                    mCell.setCellStyle(monthStyle);
                }
                // 在职天数
                row.createCell(baseC + 1).setCellValue(or.tenureDays);
                // 总在职天数
                row.createCell(baseC + 2).setCellValue(totalDaysMap.get(or));
            }

            // 合并「总在职天数」列：按 (姓名+身份证) 连续块合并 + 垂直居中
            if (mergeTotalDays && idxEmp != null && idxId != null) {
                int totalCol = outCols.size() - 1; // 最后一列
                List<CellRangeAddress> merges = new ArrayList<>();
                int n = outRows.size();
                if (n > 0) {
                    int blockStart = 0;
                    String curKey = personKey(strVal(outRows.get(0).cells.get(COL_EMP)),
                            strVal(outRows.get(0).cells.get(COL_ID)));
                    for (int i = 1; i <= n; i++) {
                        String k = (i < n) ? personKey(strVal(outRows.get(i).cells.get(COL_EMP)),
                                strVal(outRows.get(i).cells.get(COL_ID))) : null;
                        if (i == n || !k.equals(curKey)) {
                            int blockEnd = i - 1;
                            if (blockEnd > blockStart) {
                                // POI 行号 = outRows 下标 + 1（0 行为表头）
                                merges.add(new CellRangeAddress(blockStart + 1, blockEnd + 1, totalCol, totalCol));
                            }
                            if (i < n) {
                                blockStart = i;
                                curKey = k;
                            }
                        }
                    }
                }
                CellStyle center = wb.createCellStyle();
                center.setVerticalAlignment(VerticalAlignment.CENTER);
                center.setAlignment(HorizontalAlignment.CENTER);
                for (CellRangeAddress mr : merges) {
                    sheet.addMergedRegion(mr);
                    for (int rr = mr.getFirstRow(); rr <= mr.getLastRow(); rr++) {
                        Cell c = sheet.getRow(rr).getCell(totalCol);
                        c.setCellStyle(center);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                wb.write(fos);
            }
        }
    }

    private void writeOriginalCell(Cell cell, Object val, CellStyle dateStyle) {
        if (val == null) {
            cell.setBlank();
            return;
        }
        if (val instanceof LocalDate) {
            cell.setCellValue((LocalDate) val);
            cell.setCellStyle(dateStyle);
        } else if (val instanceof Double) {
            cell.setCellValue((Double) val);
        } else if (val instanceof Boolean) {
            cell.setCellValue((Boolean) val);
        } else {
            // 身份证等文本列：以字符串写入，避免 18 位科学计数法
            cell.setCellValue(String.valueOf(val));
        }
    }

    // ============================ 读取单元（保留类型） ============================
    private Object readCellValue(Cell cell, String colName) {
        if (cell == null) return "";
        if (COL_ID.equals(colName)) return getTextSafe(cell); // 身份证：强制文本安全
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getLocalDateTimeCellValue().toLocalDate();
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    CellType ct = cell.getCachedFormulaResultType();
                    if (ct == CellType.NUMERIC) {
                        if (DateUtil.isCellDateFormatted(cell)) return cell.getLocalDateTimeCellValue().toLocalDate();
                        return cell.getNumericCellValue();
                    } else if (ct == CellType.STRING) {
                        return cell.getStringCellValue();
                    } else if (ct == CellType.BOOLEAN) {
                        return cell.getBooleanCellValue();
                    }
                    return "";
                } catch (Exception ex) {
                    return "";
                }
            default:
                return "";
        }
    }

    /** 身份证等需要以文本保留的列：数字也转成无科学计数法的字符串 */
    private String getTextSafe(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate().format(DATE_FMT);
                }
                return new BigDecimal(cell.getNumericCellValue()).toPlainString();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    CellType ct = cell.getCachedFormulaResultType();
                    if (ct == CellType.NUMERIC) {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            return cell.getLocalDateTimeCellValue().toLocalDate().format(DATE_FMT);
                        }
                        return new BigDecimal(cell.getNumericCellValue()).toPlainString();
                    } else if (ct == CellType.STRING) {
                        return cell.getStringCellValue();
                    }
                    return "";
                } catch (Exception e) {
                    return "";
                }
            default:
                return "";
        }
    }

    // ============================ 日期解析工具 ============================
    private LocalDate parseDate(String raw) {
        if (raw == null) return null;
        String s = raw.trim();
        for (DateTimeFormatter f : DATE_PARSERS) {
            try {
                return LocalDate.parse(s, f);
            } catch (DateTimeParseException ignored) {
            }
        }
        return null;
    }

    private LocalDate parseDateValue(Object val) {
        if (val == null) return null;
        if (val instanceof LocalDate) return (LocalDate) val;
        if (val instanceof String) return parseDate((String) val);
        return null;
    }

    private String strVal(Object val) {
        if (val == null) return "";
        if (val instanceof LocalDate) return ((LocalDate) val).format(DATE_FMT);
        return String.valueOf(val);
    }

    private String personKey(String name, String id) {
        return (name == null ? "" : name) + "\u0001" + (id == null ? "" : id);
    }

    private LocalDate minDate(LocalDate a, LocalDate b) {
        return a.isBefore(b) ? a : b;
    }

    private LocalDate maxDate(LocalDate a, LocalDate b) {
        return a.isAfter(b) ? a : b;
    }

    // ============================ 核心算法（1:1 移植自 Python） ============================

    /** 闭区间 [start, end] 按 (年,月) 聚合的天数 */
    private Map<YM, Integer> daysByMonth(LocalDate start, LocalDate end) {
        Map<YM, Integer> res = new LinkedHashMap<>();
        if (start.isAfter(end)) return res;
        int y = start.getYear(), m = start.getMonthValue();
        while (y < end.getYear() || (y == end.getYear() && m <= end.getMonthValue())) {
            int dim = LocalDate.of(y, m, 1).lengthOfMonth();
            LocalDate segStart = LocalDate.of(y, m, 1);
            LocalDate segEnd = LocalDate.of(y, m, dim);
            LocalDate s = maxDate(segStart, start);
            LocalDate e = minDate(segEnd, end);
            int days = (int) ChronoUnit.DAYS.between(s, e) + 1;
            res.put(new YM(y, m), days);
            if (m == 12) {
                y++;
                m = 1;
            } else {
                m++;
            }
        }
        return res;
    }

    /** 从 [start, end] 中减去 [subStart, subEnd]，返回剩余闭区间列表 */
    private List<Interval> subtractInterval(LocalDate start, LocalDate end, LocalDate subStart, LocalDate subEnd) {
        if (subEnd.isBefore(start) || subStart.isAfter(end)) {
            return Collections.singletonList(new Interval(start, end));
        }
        List<Interval> res = new ArrayList<>();
        if (subStart.isAfter(start)) {
            res.add(new Interval(start, subStart.minusDays(1)));
        }
        if (subEnd.isBefore(end)) {
            res.add(new Interval(subEnd.plusDays(1), end));
        }
        return res;
    }

    /** 合并重叠或相邻的闭区间 */
    private List<Interval> mergeIntervals(List<Interval> intervals) {
        if (intervals.isEmpty()) return new ArrayList<>();
        List<Interval> sorted = new ArrayList<>(intervals);
        sorted.sort(Comparator.comparing(i -> i.start));
        List<Interval> merged = new ArrayList<>();
        merged.add(sorted.get(0));
        for (int i = 1; i < sorted.size(); i++) {
            Interval cur = sorted.get(i);
            Interval last = merged.get(merged.size() - 1);
            if (!cur.start.isAfter(last.end.plusDays(1))) {
                merged.set(merged.size() - 1, new Interval(last.start, maxDate(last.end, cur.end)));
            } else {
                merged.add(cur);
            }
        }
        return merged;
    }

    /** 根据 includeHireDay 返回区间起算日 */
    private LocalDate hireStartDate(LocalDate hire, boolean includeHireDay) {
        return includeHireDay ? hire : hire.plusDays(1);
    }

    /** 计算某个人主在职区间扣除短离职后的年月天数 */
    private Map<YM, Integer> computeMainEmployment(List<LocalDate> employedHires,
                                                   List<Interval> leaveIntervals,
                                                   LocalDate statEnd, LocalDate asOf,
                                                   boolean includeHireDay) {
        if (employedHires.isEmpty()) return new LinkedHashMap<>();

        LocalDate earliestHire = employedHires.stream().min(Comparator.naturalOrder()).get();
        LocalDate baseStart;
        if (!leaveIntervals.isEmpty()) {
            LocalDate earliestLeaveHire = leaveIntervals.stream()
                    .map(i -> i.start).min(Comparator.naturalOrder()).get();
            baseStart = maxDate(earliestHire, earliestLeaveHire);
        } else {
            baseStart = earliestHire;
        }

        LocalDate mainStart = hireStartDate(baseStart, includeHireDay);
        LocalDate mainEnd = minDate(statEnd, asOf);

        if (mainStart.isAfter(mainEnd)) return new LinkedHashMap<>();

        List<Interval> remaining = new ArrayList<>();
        remaining.add(new Interval(mainStart, mainEnd));

        List<Interval> mergedLeaves = mergeIntervals(leaveIntervals);
        for (Interval lv : mergedLeaves) {
            if (lv.end.isBefore(mainStart) || lv.start.isAfter(mainEnd)) continue;
            List<Interval> newRemaining = new ArrayList<>();
            for (Interval rem : remaining) {
                newRemaining.addAll(subtractInterval(rem.start, rem.end, lv.start, lv.end));
            }
            remaining = newRemaining;
            if (remaining.isEmpty()) break;
        }

        Map<YM, Integer> result = new LinkedHashMap<>();
        for (Interval rem : remaining) {
            Map<YM, Integer> d = daysByMonth(rem.start, rem.end);
            for (Map.Entry<YM, Integer> en : d.entrySet()) {
                result.merge(en.getKey(), en.getValue(), Integer::sum);
            }
        }
        return result;
    }

    /** 单条离职行的短离职闭区间 [start, end] */
    private Interval computeLeaveInterval(LocalDate hire, LocalDate depart, boolean includeHireDay) {
        return new Interval(hireStartDate(hire, includeHireDay), depart);
    }

    // ============================ 内部数据类 ============================
    static class RawRow {
        int index;
        Map<String, Object> cells = new LinkedHashMap<>();
        String status, name, id;
        LocalDate hire, depart;
    }

    static class OutRow {
        Map<String, Object> cells;
        LocalDate tenureMonth; // null => 空
        int tenureDays;

        OutRow(RawRow src) {
            this.cells = new LinkedHashMap<>(src.cells);
        }
    }

    static class Interval {
        LocalDate start, end;

        Interval(LocalDate start, LocalDate end) {
            this.start = start;
            this.end = end;
        }
    }

    static class LeaveRec {
        RawRow row;
        LocalDate start, end;

        LeaveRec(RawRow row, LocalDate start, LocalDate end) {
            this.row = row;
            this.start = start;
            this.end = end;
        }
    }

    static class YM {
        int year, month;

        YM(int year, int month) {
            this.year = year;
            this.month = month;
        }

        long key() {
            return (long) year * 12 + month;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof YM)) return false;
            YM ym = (YM) o;
            return year == ym.year && month == ym.month;
        }

        @Override
        public int hashCode() {
            return year * 31 + month;
        }
    }
}
