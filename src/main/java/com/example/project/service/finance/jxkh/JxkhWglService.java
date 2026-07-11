package com.example.project.service.finance.jxkh;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 绩效考核 - 稳岗率分析（cw_wgl_zhf.py + frame_wgl.py）
 * 与现有 CwWglZhfService 的差异：
 * 1. 增加"总人数"列过滤
 * 2. 增加工厂实际结算有效性过滤
 * 3. 稳岗率档位计算消除浮点误差
 * 4. 保留原始工作表并追加新工作表
 */
@Service
public class JxkhWglService {

    private static final DateTimeFormatter MONTH_FMT = DateTimeFormatter.ofPattern("yyyy/MM");
    private static final DateTimeFormatter[] DATE_PARSERS = {
        DateTimeFormatter.ofPattern("yyyy/MM/dd"),
        DateTimeFormatter.ofPattern("yyyy-MM-dd"),
        DateTimeFormatter.ofPattern("yyyy/MM"),
        DateTimeFormatter.ofPattern("yyyy-MM"),
        DateTimeFormatter.ofPattern("yyyy年MM月dd日"),
        DateTimeFormatter.ofPattern("yyyyMMdd")
    };

    static class EmpRecord {
        String name;
        String center;
        String status;
        String entryMonth;
        String exitMonth;
        LocalDate entryDate;
        LocalDate exitDate;
        String projectSite;   // 项目驻场
        String projectManager; // 项目经理
        String actualSettlement;
        boolean validSettlement;
        String totalNumber;    // 总人数（用于 buildRefresh 内部过滤）
    }

    public static class AnalyzeResult {
        public String filePath;
        public List<String> logs = new ArrayList<>();
    }

    public AnalyzeResult analyze(MultipartFile file, String sheetName, String startMonth, String endMonth,
                          boolean copyOriginal, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        try {
            return doAnalyze(tempFile, file.getOriginalFilename(), sheetName, startMonth, endMonth, copyOriginal, uploadDir);
        } finally {
            tempFile.delete();
        }
    }

    private AnalyzeResult doAnalyze(File inputFile, String originalFilename,
                             String sheetName, String startMonth, String endMonth,
                             boolean copyOriginal, String uploadDir) throws Exception {
        List<String> logs = new ArrayList<>();
        AnalyzeResult result = new AnalyzeResult();
        result.logs = logs;

        // 中间计算结果（读取完成后仍需使用）
        List<EmpRecord> records = null;
        RefreshData zcRefresh = null, pmRefresh = null;
        StabilitySheet zcStability = null, pmStability = null;
        List<Map<String, Object>> reward = null;
        List<Map<String, Object>> baseRate = null;
        String outPath = null;

        // 仅打开一次 Workbook，读取阶段与写出阶段（保留原表时）复用同一实例，
        // 避免同一文件被 POI 全量加载两次导致内存峰值翻倍、触发 OOM
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook wb = WorkbookFactory.create(fis)) {

            ReadResult readResult = readRecords(wb, sheetName, logs);
            records = readResult.records;
            if (records.isEmpty()) throw new Exception("没有读取到数据");
            logs.add("原始数据行数：" + records.size());
            logs.add("读取到的列名：" + String.join(", ", readResult.columnNames));

            Set<String> allMonthsSet = new TreeSet<>();
            for (EmpRecord r : records) {
                if (r.entryMonth != null) allMonthsSet.add(r.entryMonth);
                if (r.exitMonth != null) allMonthsSet.add(r.exitMonth);
            }
            List<String> allMonths = new ArrayList<>(allMonthsSet);

            final String effectiveStartMonth = (startMonth == null || startMonth.isEmpty()) ? allMonths.get(0) : startMonth;
            final String effectiveEndMonth = (endMonth == null || endMonth.isEmpty()) ? allMonths.get(allMonths.size() - 1) : endMonth;

            List<String> targetMonths = allMonths.stream()
                    .filter(m -> m.compareTo(effectiveStartMonth) >= 0 && m.compareTo(effectiveEndMonth) <= 0)
                    .collect(Collectors.toList());

            String targetMonth = effectiveEndMonth;
            logs.add("时间范围：" + effectiveStartMonth + " 至 " + effectiveEndMonth);
            logs.add("核算月份：" + targetMonth);

            zcRefresh = buildRefresh(records, "项目驻场", targetMonths, effectiveStartMonth, effectiveEndMonth, logs);
            pmRefresh = buildRefresh(records, "项目经理", targetMonths, effectiveStartMonth, effectiveEndMonth, logs);

            zcStability = buildStabilitySheet(records, targetMonth, zcRefresh, true, logs);
            pmStability = buildStabilitySheet(records, targetMonth, pmRefresh, false, logs);

            reward = buildReward(zcStability, pmStability);

            double zcBaseRate = Math.max(zcStability.avgRate, zcStability.medianRate);
            double pmBaseRate = Math.max(pmStability.avgRate, pmStability.medianRate);
            baseRate = buildBaseRate(zcBaseRate, pmBaseRate);

            logs.add("项目驻场基准数：" + formatPct(zcBaseRate));
            logs.add("项目经理基准数：" + formatPct(pmBaseRate));

            if (copyOriginal) {
                // 复用已加载的 Workbook 追加新工作表，不再重新打开文件（使用原始文件名，不含 UUID 前缀）
                String baseName = originalFilename.replaceAll("\\.xlsx$", "").replaceAll("\\.xls$", "");
                File outFile = new File(uploadDir, baseName + "_处理结果.xlsx");
                copyAndAppend(wb, outFile, zcRefresh, pmRefresh, zcStability, pmStability, reward, baseRate, logs);
                outPath = outFile.getAbsolutePath();
            }
        }  // Workbook 在此自动关闭，内存释放

        if (!copyOriginal) {
            // 不保留原表：新建 Workbook，原文件 Workbook 已关闭，避免两者内存叠加
            File outFile = new File(uploadDir, "稳岗率计算结果.xlsx");
            writeNew(outFile, zcRefresh, pmRefresh, zcStability, pmStability, reward, baseRate);
            outPath = outFile.getAbsolutePath();
        }

        result.filePath = outPath;
        return result;
    }

    static class ReadResult {
        List<EmpRecord> records = new ArrayList<>();
        List<String> columnNames = new ArrayList<>();
    }

    // ========== 读取数据（含总人数过滤） ==========
    private ReadResult readRecords(Workbook wb, String sheetName, List<String> logs) throws Exception {
        ReadResult result = new ReadResult();
        Sheet sheet = sheetName != null && !sheetName.isEmpty() ? wb.getSheet(sheetName) : wb.getSheetAt(0);
            if (sheet == null) throw new Exception("工作表不存在：" + sheetName);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new Exception("没有表头行");

            Map<String, Integer> colMap = new HashMap<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                String h = ExcelUtils.getCellStringValue(headerRow.getCell(c)).trim();
                colMap.put(h, c);
                if (!h.isEmpty()) result.columnNames.add(h);
            }

            int idxName = colMap.getOrDefault("姓名", -1);
            int idxCenter = colMap.getOrDefault("所属运营中心", -1);
            int idxStatus = colMap.getOrDefault("状态", -1);
            int idxEntryMonth = colMap.getOrDefault("当月入职", -1);
            int idxExitMonth = colMap.getOrDefault("当月离职", -1);
            int idxEntryDate = colMap.getOrDefault("入职日期", -1);
            int idxExitDate = colMap.getOrDefault("离职日期", -1);
            int idxPm = colMap.getOrDefault("项目驻场", -1);
            int idxPl = colMap.getOrDefault("项目经理", -1);

            // 查找工厂实际结算列
            int idxSettlement = -1;
            for (Map.Entry<String, Integer> e : colMap.entrySet()) {
                if (e.getKey().contains("工厂实际结算")) { idxSettlement = e.getValue(); break; }
            }
            // 查找总人数列（jxkh新增过滤）——兼容列名中间有换行符等情况
            int idxTotal = -1;
            for (Map.Entry<String, Integer> e : colMap.entrySet()) {
                String clean = e.getKey().replaceAll("\\s+", "");
                if (clean.contains("总人数")) { idxTotal = e.getValue(); break; }
            }
            if (idxTotal == -1) {
                logs.add("警告：未匹配到'总人数'列，已读取的列名：" + String.join(" | ", result.columnNames));
            } else {
                logs.add("已匹配到'总人数'列，列索引：" + idxTotal);
            }

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                EmpRecord rec = new EmpRecord();
                rec.name = getCellStr(row, idxName);
                rec.center = getCellStr(row, idxCenter);
                rec.status = getCellStr(row, idxStatus);
                rec.entryMonth = normalizeMonth(getCellStr(row, idxEntryMonth));
                rec.exitMonth = normalizeMonth(getCellStr(row, idxExitMonth));
                rec.entryDate = parseDate(getCellStr(row, idxEntryDate));
                rec.exitDate = parseDate(getCellStr(row, idxExitDate));
                rec.projectSite = getCellStr(row, idxPm);
                rec.projectManager = getCellStr(row, idxPl);
                rec.actualSettlement = getCellStr(row, idxSettlement);
                rec.validSettlement = isValidSettlementValue(rec.actualSettlement);
                // 读取"总人数"列（用于 buildRefresh 内部过滤，与 Python 原版一致）
                rec.totalNumber = getCellStr(row, idxTotal);

                result.records.add(rec);
            }
        return result;
    }

    private String getCellStr(Row row, int idx) {
        if (idx < 0) return null;
        String val = ExcelUtils.getCellStringValue(row.getCell(idx));
        if (val == null || val.trim().isEmpty() || "nan".equalsIgnoreCase(val.trim())) return null;
        return val.trim();
    }

    private String normalizeMonth(String raw) {
        if (raw == null) return null;
        for (DateTimeFormatter fmt : DATE_PARSERS) {
            try {
                LocalDate d = LocalDate.parse(raw, fmt);
                return d.format(MONTH_FMT);
            } catch (DateTimeParseException ignored) {}
        }
        if (raw.matches("\\d{4}/\\d{2}")) return raw;
        return null;
    }

    private LocalDate parseDate(String raw) {
        if (raw == null) return null;
        for (DateTimeFormatter fmt : DATE_PARSERS) {
            try { return LocalDate.parse(raw, fmt); } catch (Exception ignored) {}
        }
        return null;
    }

    private boolean isValidSettlementValue(String val) {
        if (val == null || val.isEmpty()) return false;
        return val.length() <= 50;
    }

    private boolean isValidTotalNumber(String val) {
        if (val == null || val.isEmpty()) return false;
        try { Double.parseDouble(val); return true; } catch (Exception e) { return false; }
    }

    // ========== 刷新表 ==========
    static class RefreshData {
        List<String> actors;
        List<String> months;
        String startMonth, endMonth;
        Map<String, Map<String, Integer>> entryByActor;
        Map<String, Map<String, Integer>> exitByActor;
    }

    /**
     * 构建刷新表（当月入职/离职统计）
     * 与 Python 原版 build_refresh 对应：
     * - 需要先过滤"总人数"有效行（Python 第127-132行）
     * - 再统计入职/离职人数
     * 注意：Python 原版只在 build_refresh 内部过滤"总人数"，
     *       其他函数（如 build_stability_sheet）使用未过滤的原始数据。
     *       为了严格一致，这里在 buildRefresh 内部过滤，而 readRecords 不过滤。
     */
    private RefreshData buildRefresh(List<EmpRecord> records, String roleCol,
                                      List<String> targetMonths, String startMonth, String endMonth,
                                      List<String> logs) {
        RefreshData data = new RefreshData();
        data.months = targetMonths;
        data.startMonth = startMonth;
        data.endMonth = endMonth;

        // 第一重过滤：筛选"总人数"列为有效值的行（与 Python 原版一致）
        List<EmpRecord> filteredRecords = new ArrayList<>();
        for (EmpRecord r : records) {
            if (isValidSettlementValue(r.totalNumber)) {
                filteredRecords.add(r);
            }
        }
        // 使用过滤后的记录进行后续统计
        List<EmpRecord> useRecords = filteredRecords.isEmpty() ? records : filteredRecords;
        if (!filteredRecords.isEmpty()) {
            logs.add("[" + roleCol + "] 总人数过滤前：" + records.size() + " 行，过滤后：" + filteredRecords.size() + " 行");
        } else {
            logs.add("[" + roleCol + "] 警告：未找到包含'总人数'的有效列，使用所有数据");
        }

        Set<String> actorSet = new TreeSet<>();
        for (EmpRecord r : useRecords) {
            String actor = roleCol.equals("项目驻场") ? r.projectSite : r.projectManager;
            if (actor != null && !actor.isEmpty()) actorSet.add(actor);
        }
        data.actors = new ArrayList<>(actorSet);

        data.entryByActor = new HashMap<>();
        data.exitByActor = new HashMap<>();
        for (String a : data.actors) {
            data.entryByActor.put(a, new HashMap<>());
            data.exitByActor.put(a, new HashMap<>());
            for (String m : targetMonths) {
                data.entryByActor.get(a).put(m, 0);
                data.exitByActor.get(a).put(m, 0);
            }
        }

        for (EmpRecord r : useRecords) {
            String actor = roleCol.equals("项目驻场") ? r.projectSite : r.projectManager;
            if (actor == null || !actorSet.contains(actor)) continue;
            if (r.entryMonth != null && targetMonths.contains(r.entryMonth)) {
                data.entryByActor.get(actor).merge(r.entryMonth, 1, Integer::sum);
            }
            if (r.exitMonth != null && targetMonths.contains(r.exitMonth)) {
                data.exitByActor.get(actor).merge(r.exitMonth, 1, Integer::sum);
            }
        }
        return data;
    }

    // ========== 稳岗率表 ==========
    static class StabilityRow {
        String name;
        int monthEndCount;
        int totalOnJobCount;
        double stabilityRate;
        int entryCount;
        double unitPrice;
        double entrySubsidy;
        int settlementCount;
        double adjustedUnitPrice;
        double stabilitySubsidy;
    }

    static class StabilitySheet {
        List<StabilityRow> rows;
        double avgRate;
        double medianRate;
        double baseRate;
    }

    private StabilitySheet buildStabilitySheet(List<EmpRecord> records, String endMonth,
                                                RefreshData refresh, boolean isProjectSite,
                                                List<String> logs) {
        StabilitySheet sheet = new StabilitySheet();
        String roleLabel = isProjectSite ? "项目驻场" : "项目经理";
        // isProjectSite=true: 项目驻场(单价5元); false: 项目经理(单价1元)
        double defaultUnitPrice = isProjectSite ? 5.0 : 1.0;

        List<String> exitMonths = refresh.months.stream()
                .filter(m -> m.compareTo(endMonth) < 0).collect(Collectors.toList());

        // 统计当月结算总人数（需过滤 validSettlement）
        Map<String, Integer> settlementCounts = new HashMap<>();
        for (EmpRecord r : records) {
            String actor = isProjectSite ? r.projectSite : r.projectManager;
            if (actor == null || actor.isEmpty()) continue;
            if (r.validSettlement) settlementCounts.merge(actor, 1, Integer::sum);
        }

        // 统计当月入职人数（需过滤 validSettlement，与 Python 原版一致）
        Map<String, Integer> entryCounts = new HashMap<>();
        for (EmpRecord r : records) {
            String actor = isProjectSite ? r.projectSite : r.projectManager;
            if (actor == null || actor.isEmpty()) continue;
            if (!r.validSettlement) continue;
            if (r.entryMonth != null && r.entryMonth.equals(endMonth)) {
                entryCounts.merge(actor, 1, Integer::sum);
            }
        }

        logs.add(roleLabel + "原始数据行数：" + records.size() + "，过滤后的有效数据行数：" + settlementCounts.values().stream().mapToInt(Integer::intValue).sum());

        List<StabilityRow> allRows = new ArrayList<>();
        for (String actor : refresh.actors) {
            StabilityRow row = new StabilityRow();
            row.name = actor;

            int entryTotal = refresh.entryByActor.getOrDefault(actor, Collections.emptyMap())
                    .values().stream().mapToInt(Integer::intValue).sum();
            int exitTotal = refresh.exitByActor.getOrDefault(actor, Collections.emptyMap())
                    .values().stream().mapToInt(Integer::intValue).sum();
            int exitUntilPrev = 0;
            Map<String, Integer> exitMap = refresh.exitByActor.getOrDefault(actor, Collections.emptyMap());
            for (String m : exitMonths) exitUntilPrev += exitMap.getOrDefault(m, 0);

            row.monthEndCount = entryTotal - exitTotal;
            row.totalOnJobCount = entryTotal - exitUntilPrev;
            row.stabilityRate = row.totalOnJobCount > 0
                    ? roundHalfUp((double) row.monthEndCount / row.totalOnJobCount, 2) : 0.0;

            // 修复：当月入职人数需要过滤 validSettlement（与 Python 原版一致）
            row.entryCount = entryCounts.getOrDefault(actor, 0);
            row.unitPrice = actor.contains("锦途") ? 0.0 : defaultUnitPrice;
            row.entrySubsidy = row.entryCount * row.unitPrice;
            row.settlementCount = settlementCounts.getOrDefault(actor, 0);
            allRows.add(row);
        }

        List<StabilityRow> jintuRows = allRows.stream().filter(r -> r.name.contains("锦途")).collect(Collectors.toList());
        List<StabilityRow> otherRows = allRows.stream().filter(r -> !r.name.contains("锦途"))
                .sorted(Comparator.comparingDouble(r -> r.stabilityRate)).collect(Collectors.toList());

        // 记录各角色人数统计
        StringBuilder statSb = new StringBuilder();
        statSb.append("各").append(roleLabel).append("人数统计：{");
        boolean first = true;
        for (StabilityRow sr : allRows) {
            if (sr.name == null || sr.name.contains("锦途")) continue;
            if (!first) statSb.append(", ");
            statSb.append("'").append(sr.name).append("': ").append(sr.settlementCount);
            first = false;
        }
        statSb.append("}");
        logs.add(statSb.toString());

        double[] rates = otherRows.stream().mapToDouble(r -> r.stabilityRate).toArray();
        double avg = rates.length > 0 ? Arrays.stream(rates).average().orElse(0) : 0;
        double median = rates.length > 0 ? calcMedian(rates) : 0;
        double baseRate = roundHalfUp(Math.max(avg, median), 2);

        logs.add(roleLabel + " — 平均数：" + roundHalfUp(avg, 2) + "，中位数：" + roundHalfUp(median, 2) + "，基准数：" + baseRate);

        for (StabilityRow r : otherRows) {
            r.adjustedUnitPrice = calcAdjustedUnitPrice(r.stabilityRate, r.name, baseRate, isProjectSite);
            r.stabilitySubsidy = r.settlementCount * r.adjustedUnitPrice;
        }
        for (StabilityRow r : jintuRows) {
            r.adjustedUnitPrice = 0;
            r.stabilitySubsidy = 0;
        }

        // 拼接
        List<StabilityRow> finalRows = new ArrayList<>();
        finalRows.addAll(otherRows);
        finalRows.addAll(jintuRows);
        finalRows.add(null);

        StabilityRow totalRow = new StabilityRow();
        totalRow.name = "总计";
        totalRow.monthEndCount = sumInt(otherRows, jintuRows, r -> r.monthEndCount);
        totalRow.totalOnJobCount = sumInt(otherRows, jintuRows, r -> r.totalOnJobCount);
        totalRow.stabilityRate = totalRow.totalOnJobCount > 0
                ? roundHalfUp((double) totalRow.monthEndCount / totalRow.totalOnJobCount, 2) : 0.0;
        totalRow.entryCount = sumInt(otherRows, jintuRows, r -> r.entryCount);
        totalRow.unitPrice = sumDbl(otherRows, jintuRows, r -> r.unitPrice);
        totalRow.entrySubsidy = sumDbl(otherRows, jintuRows, r -> r.entrySubsidy);
        totalRow.settlementCount = sumInt(otherRows, jintuRows, r -> r.settlementCount);
        totalRow.adjustedUnitPrice = sumDbl(otherRows, jintuRows, r -> r.adjustedUnitPrice);
        totalRow.stabilitySubsidy = sumDbl(otherRows, jintuRows, r -> r.stabilitySubsidy);
        finalRows.add(totalRow);
        finalRows.add(null);

        StabilityRow avgRow = new StabilityRow(); avgRow.name = "平均数"; avgRow.stabilityRate = roundHalfUp(avg, 2);
        finalRows.add(avgRow);
        StabilityRow medianRow = new StabilityRow(); medianRow.name = "中位数"; medianRow.stabilityRate = roundHalfUp(median, 2);
        finalRows.add(medianRow);

        sheet.rows = finalRows;
        sheet.avgRate = avg;
        sheet.medianRate = median;
        sheet.baseRate = baseRate;
        return sheet;
    }

    private double calcMedian(double[] arr) {
        double[] sorted = arr.clone();
        Arrays.sort(sorted);
        int n = sorted.length;
        return n % 2 == 0 ? (sorted[n/2 - 1] + sorted[n/2]) / 2 : sorted[n/2];
    }

    private int sumInt(List<StabilityRow> a, List<StabilityRow> b, java.util.function.Function<StabilityRow, Integer> fn) {
        return a.stream().mapToInt(fn::apply).sum() + b.stream().mapToInt(fn::apply).sum();
    }
    private double sumDbl(List<StabilityRow> a, List<StabilityRow> b, java.util.function.Function<StabilityRow, Double> fn) {
        return a.stream().mapToDouble(fn::apply).sum() + b.stream().mapToDouble(fn::apply).sum();
    }

    /**
     * 计算折算后单价
     * @param isProjectSite true=项目驻场(单价范围0~10元); false=项目经理(单价范围0~2元)
     */
    private double calcAdjustedUnitPrice(double rate, String name, double baseRate, boolean isProjectSite) {
        if (name.contains("锦途")) return 0.0;
        double diff = rate - baseRate;
        // 消除浮点误差：math.floor(round(diff / 0.05, 10))
        double gearD = Math.round(diff / 0.05 * 1e10) / 1e10;
        int gear = (int) Math.floor(gearD);

        if (isProjectSite) {
            // 项目驻场：单价 0~10 元
            if (gear >= 6) return 10.0;
            if (gear < -5) return 0.0;
            return gear * 0.5 + 5;
        } else {
            // 项目经理：单价 0~2 元
            if (gear >= 6) return 2.0;
            if (gear < -5) return 0.0;
            return gear * 0.1 + 1;
        }
    }

    // ========== 奖励汇总 ==========
    private List<Map<String, Object>> buildReward(StabilitySheet zc, StabilitySheet pm) {
        Map<String, double[]> zcMap = new LinkedHashMap<>();
        for (StabilityRow r : zc.rows) {
            if (r == null || "总计".equals(r.name) || "平均数".equals(r.name) || "中位数".equals(r.name)) continue;
            // 过滤包含"对比结果"的行（与 Python 原版一致）
            if (r.name != null && r.name.contains("对比结果")) continue;
            zcMap.put(r.name, new double[]{r.entrySubsidy, r.stabilitySubsidy});
        }
        Map<String, double[]> pmMap = new LinkedHashMap<>();
        for (StabilityRow r : pm.rows) {
            if (r == null || "总计".equals(r.name) || "平均数".equals(r.name) || "中位数".equals(r.name)) continue;
            // 过滤包含"这是对比的结果"的行（与 Python 原版一致）
            if (r.name != null && r.name.contains("这是对比的结果")) continue;
            pmMap.put(r.name, new double[]{r.entrySubsidy, r.stabilitySubsidy});
        }

        Set<String> allNames = new TreeSet<>(zcMap.keySet());
        allNames.addAll(pmMap.keySet());

        List<Map<String, Object>> result = new ArrayList<>();
        double totalZcEntry = 0, totalZcStable = 0, totalPmEntry = 0, totalPmStable = 0;

        for (String name : allNames) {
            Map<String, Object> row = new LinkedHashMap<>();
            row.put("姓名", name);
            double[] zcData = zcMap.getOrDefault(name, new double[]{0, 0});
            double[] pmData = pmMap.getOrDefault(name, new double[]{0, 0});
            row.put("入职补贴（5元/人）", zcData[0]);
            row.put("稳岗补贴_驻场", zcData[1]);
            row.put("入职补贴（1元/人）", pmData[0]);
            row.put("稳岗补贴_经理", pmData[1]);
            double sum = zcData[0] + zcData[1] + pmData[0] + pmData[1];
            row.put("合计", roundHalfUp(sum, 2));
            row.put("备注", "");
            result.add(row);
            totalZcEntry += zcData[0]; totalZcStable += zcData[1];
            totalPmEntry += pmData[0]; totalPmStable += pmData[1];
        }

        Map<String, Object> totalRow = new LinkedHashMap<>();
        totalRow.put("姓名", "总计");
        totalRow.put("入职补贴（5元/人）", totalZcEntry);
        totalRow.put("稳岗补贴_驻场", totalZcStable);
        totalRow.put("入职补贴（1元/人）", totalPmEntry);
        totalRow.put("稳岗补贴_经理", totalPmStable);
        totalRow.put("合计", roundHalfUp(totalZcEntry + totalZcStable + totalPmEntry + totalPmStable, 2));
        totalRow.put("备注", "");
        result.add(totalRow);
        return result;
    }

    // ========== 基准数 ==========
    private List<Map<String, Object>> buildBaseRate(double zcBaseRate, double pmBaseRate) {
        List<Map<String, Object>> rows = new ArrayList<>();
        Map<String, Object> titleRow = new LinkedHashMap<>();
        titleRow.put("col2", "项目驻场"); titleRow.put("col9", "项目经理");
        rows.add(titleRow);

        for (int gear = 10; gear >= -8; gear--) {
            Map<String, Object> row = new LinkedHashMap<>();
            double rateZc = zcBaseRate + gear * 0.05;
            double unitZc; String noteZc = "";
            if (gear >= 6) { unitZc = 10.0; noteZc = gear == 6 ? ">=8元时，实际按照10元结算" : ""; }
            else if (gear < -5) { unitZc = 0.0; noteZc = gear == -6 ? "<2元时，不结算" : ""; }
            else unitZc = gear * 0.5 + 5;

            row.put("col1", gear == 0 ? roundHalfUp(zcBaseRate, 2) : "");
            row.put("col2", gear);
            row.put("col3", formatPct(rateZc));
            row.put("col4", unitZc);
            row.put("col5", noteZc);
            row.put("col6", "");

            double ratePm = pmBaseRate + gear * 0.05;
            double unitPm; String notePm = "";
            if (gear >= 6) { unitPm = 2.0; notePm = gear == 6 ? ">=1.6元时，实际按照2元结算" : ""; }
            else if (gear < -5) { unitPm = 0.0; notePm = gear == -6 ? "<0.4元时，不结算" : ""; }
            else unitPm = gear * 0.1 + 1;

            row.put("col8", "");
            row.put("col9", gear == 0 ? roundHalfUp(pmBaseRate, 2) : "");
            row.put("col10", gear);
            row.put("col11", formatPct(ratePm));
            row.put("col12", unitPm);
            row.put("col13", notePm);
            rows.add(row);
        }
        return rows;
    }

    private String formatPct(double rate) {
        return (int) Math.round(rate * 100) + "%";
    }

    // ========== 写出（保留原表） ==========
    private void copyAndAppend(Workbook wb, File outputFile,
                                RefreshData zcRefresh, RefreshData pmRefresh,
                                StabilitySheet zcStability, StabilitySheet pmStability,
                                List<Map<String, Object>> reward,
                                List<Map<String, Object>> baseRate,
                                List<String> logs) throws Exception {
        // 复用已在 doAnalyze 中打开的 Workbook（避免二次全量加载导致 OOM）

        // 仅保留指定的原始工作表，移除其余工作表
        List<String> removedSheets = new ArrayList<>();
            for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
                String sheetName = wb.getSheetName(i);
                if (!shouldKeepSheet(sheetName)) {
                    wb.removeSheetAt(i);
                    removedSheets.add(sheetName);
                }
            }
            if (!removedSheets.isEmpty()) {
                logs.add("已移除原始工作表：" + String.join(", ", removedSheets));
            }

            deleteIfExists(wb, "项目驻场");
            deleteIfExists(wb, "项目经理");
            deleteIfExists(wb, "稳岗率");
            deleteIfExists(wb, "驻场&项目经理 - 奖励汇总");
            deleteIfExists(wb, "基准数");

            writeRefreshSheet(wb, "项目驻场", zcRefresh);
            writeRefreshSheet(wb, "项目经理", pmRefresh);
            writeStabilityCombinedSheet(wb, "稳岗率", zcStability, pmStability);
            writeRewardSheet(wb, "驻场&项目经理 - 奖励汇总", reward);
            writeBaseRateSheet(wb, "基准数", baseRate);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
    }

    private void deleteIfExists(Workbook wb, String name) {
        int idx = wb.getSheetIndex(name);
        if (idx >= 0) wb.removeSheetAt(idx);
    }

    /**
     * 判断原始工作表是否需要保留
     * 保留规则：
     * 1. 精确匹配："运营中心-项目经理"、"员工信息汇总"
     * 2. 前缀匹配："代工费-"、"百事通报备名单-"、"员工信息汇总-"
     */
    private boolean shouldKeepSheet(String name) {
        if (name == null) return false;
        if ("运营中心-项目经理".equals(name)) return true;
        if ("员工信息汇总".equals(name)) return true;
        if (name.startsWith("代工费-")) return true;
        if (name.startsWith("百事通报备名单-")) return true;
        if (name.startsWith("员工信息汇总-")) return true;
        return false;
    }

    private void writeNew(File outputFile, RefreshData zcRefresh, RefreshData pmRefresh,
                          StabilitySheet zcStability, StabilitySheet pmStability,
                          List<Map<String, Object>> reward, List<Map<String, Object>> baseRate) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            writeRefreshSheet(wb, "项目驻场", zcRefresh);
            writeRefreshSheet(wb, "项目经理", pmRefresh);
            writeStabilityCombinedSheet(wb, "稳岗率", zcStability, pmStability);
            writeRewardSheet(wb, "驻场&项目经理 - 奖励汇总", reward);
            writeBaseRateSheet(wb, "基准数", baseRate);
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    private void writeRefreshSheet(Workbook wb, String title, RefreshData data) {
        Sheet sheet = wb.createSheet(title);
        int r = 0;
        List<String> exitMonths = data.months.stream()
                .filter(m -> m.compareTo(data.endMonth) < 0).collect(Collectors.toList());

        // ========== 当月入职 ==========
        Row inHeader = sheet.createRow(r++);
        inHeader.createCell(0).setCellValue("当月入职");
        for (int i = 0; i < data.actors.size(); i++) {
            inHeader.createCell(i + 1).setCellValue(data.actors.get(i));
        }
        inHeader.createCell(data.actors.size() + 1).setCellValue("总计");

        for (String m : data.months) {
            Row row = sheet.createRow(r++);
            row.createCell(0).setCellValue(m);
            int total = 0;
            for (int i = 0; i < data.actors.size(); i++) {
                String actor = data.actors.get(i);
                int count = data.entryByActor.getOrDefault(actor, Collections.emptyMap()).getOrDefault(m, 0);
                row.createCell(i + 1).setCellValue(count);
                total += count;
            }
            row.createCell(data.actors.size() + 1).setCellValue(total);
        }

        Row inTotal = sheet.createRow(r++);
        inTotal.createCell(0).setCellValue("总计");
        int grandEntryTotal = 0;
        for (int i = 0; i < data.actors.size(); i++) {
            String actor = data.actors.get(i);
            int total = data.entryByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            inTotal.createCell(i + 1).setCellValue(total);
            grandEntryTotal += total;
        }
        inTotal.createCell(data.actors.size() + 1).setCellValue(grandEntryTotal);

        // 空一行
        r++;

        // ========== 当月离职 ==========
        Row outHeader = sheet.createRow(r++);
        outHeader.createCell(0).setCellValue("当月离职");
        for (int i = 0; i < data.actors.size(); i++) {
            outHeader.createCell(i + 1).setCellValue(data.actors.get(i));
        }
        outHeader.createCell(data.actors.size() + 1).setCellValue("总计");

        for (String m : data.months) {
            Row row = sheet.createRow(r++);
            row.createCell(0).setCellValue(m);
            int total = 0;
            for (int i = 0; i < data.actors.size(); i++) {
                String actor = data.actors.get(i);
                int count = data.exitByActor.getOrDefault(actor, Collections.emptyMap()).getOrDefault(m, 0);
                row.createCell(i + 1).setCellValue(count);
                total += count;
            }
            row.createCell(data.actors.size() + 1).setCellValue(total);
        }

        Row outTotal = sheet.createRow(r++);
        outTotal.createCell(0).setCellValue("总计");
        int grandExitTotal = 0;
        for (int i = 0; i < data.actors.size(); i++) {
            String actor = data.actors.get(i);
            int total = data.exitByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            outTotal.createCell(i + 1).setCellValue(total);
            grandExitTotal += total;
        }
        outTotal.createCell(data.actors.size() + 1).setCellValue(grandExitTotal);

        // 空一行
        r++;

        // ========== 月末统计 ==========
        Row meRow = sheet.createRow(r++);
        meRow.createCell(0).setCellValue("月末在职人数");
        int meTotal = 0;
        for (int i = 0; i < data.actors.size(); i++) {
            String actor = data.actors.get(i);
            int entry = data.entryByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            int exit = data.exitByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            int val = entry - exit;
            meRow.createCell(i + 1).setCellValue(val);
            meTotal += val;
        }
        meRow.createCell(data.actors.size() + 1).setCellValue(meTotal);

        Row tjRow = sheet.createRow(r++);
        tjRow.createCell(0).setCellValue("当月总在职人数");
        int tjTotal = 0;
        for (int i = 0; i < data.actors.size(); i++) {
            String actor = data.actors.get(i);
            int entry = data.entryByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            int exitPrev = data.exitByActor.getOrDefault(actor, Collections.emptyMap()).entrySet().stream()
                    .filter(e -> exitMonths.contains(e.getKey())).mapToInt(Map.Entry::getValue).sum();
            int val = entry - exitPrev;
            tjRow.createCell(i + 1).setCellValue(val);
            tjTotal += val;
        }
        tjRow.createCell(data.actors.size() + 1).setCellValue(tjTotal);

        Row srRow = sheet.createRow(r++);
        srRow.createCell(0).setCellValue("稳岗率");
        for (int i = 0; i < data.actors.size(); i++) {
            String actor = data.actors.get(i);
            int entry = data.entryByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            int exit = data.exitByActor.getOrDefault(actor, Collections.emptyMap()).values().stream().mapToInt(Integer::intValue).sum();
            int exitPrev = data.exitByActor.getOrDefault(actor, Collections.emptyMap()).entrySet().stream()
                    .filter(e -> exitMonths.contains(e.getKey())).mapToInt(Map.Entry::getValue).sum();
            int me = entry - exit;
            int tj = entry - exitPrev;
            srRow.createCell(i + 1).setCellValue(tj > 0 ? roundHalfUp((double) me / tj, 2) : 0.0);
        }
        srRow.createCell(data.actors.size() + 1).setCellValue(tjTotal > 0 ? roundHalfUp((double) meTotal / tjTotal, 2) : 0.0);
    }

    private void writeStabilityCombinedSheet(Workbook wb, String title,
                                              StabilitySheet zc, StabilitySheet pm) {
        Sheet sheet = wb.createSheet(title);
        String[] zcCols = {"项目驻场", "月末在职人数", "当月总在职人数", "稳岗率",
                "当月入职人数", "单价", "入职补贴（5元/人）", "当月结算总人数",
                "折算后单价5元/人 (期初)-稳岗", "稳岗补贴"};
        String[] pmCols = {"项目经理", "月末在职人数", "当月总在职人数", "稳岗率",
                "当月入职人数", "单价", "入职补贴（1元/人）", "当月结算总人数",
                "折算后单价5元/人 (期初)-稳岗", "稳岗补贴"};
        int maxRows = Math.max(zc.rows.size(), pm.rows.size());

        Row header = sheet.createRow(0);
        for (int c = 0; c < zcCols.length; c++) header.createCell(c).setCellValue(zcCols[c]);
        int offset = zcCols.length + 3;
        for (int c = 0; c < pmCols.length; c++) header.createCell(offset + c).setCellValue(pmCols[c]);

        for (int r = 0; r < maxRows; r++) {
            Row row = sheet.createRow(r + 1);
            if (r < zc.rows.size()) {
                StabilityRow sr = zc.rows.get(r);
                if (sr != null) {
                    row.createCell(0).setCellValue(sr.name);
                    if ("平均数".equals(sr.name) || "中位数".equals(sr.name)) {
                        row.createCell(3).setCellValue(sr.stabilityRate);
                    } else {
                        row.createCell(1).setCellValue(sr.monthEndCount);
                        row.createCell(2).setCellValue(sr.totalOnJobCount);
                        row.createCell(3).setCellValue(sr.stabilityRate);
                        row.createCell(4).setCellValue(sr.entryCount);
                        row.createCell(5).setCellValue(sr.unitPrice);
                        row.createCell(6).setCellValue(sr.entrySubsidy);
                        row.createCell(7).setCellValue(sr.settlementCount);
                        row.createCell(8).setCellValue(sr.adjustedUnitPrice);
                        row.createCell(9).setCellValue(sr.stabilitySubsidy);
                    }
                }
            }
            if (r < pm.rows.size()) {
                StabilityRow sr = pm.rows.get(r);
                if (sr != null) {
                    row.createCell(offset + 0).setCellValue(sr.name);
                    if ("平均数".equals(sr.name) || "中位数".equals(sr.name)) {
                        row.createCell(offset + 3).setCellValue(sr.stabilityRate);
                    } else {
                        row.createCell(offset + 1).setCellValue(sr.monthEndCount);
                        row.createCell(offset + 2).setCellValue(sr.totalOnJobCount);
                        row.createCell(offset + 3).setCellValue(sr.stabilityRate);
                        row.createCell(offset + 4).setCellValue(sr.entryCount);
                        row.createCell(offset + 5).setCellValue(sr.unitPrice);
                        row.createCell(offset + 6).setCellValue(sr.entrySubsidy);
                        row.createCell(offset + 7).setCellValue(sr.settlementCount);
                        row.createCell(offset + 8).setCellValue(sr.adjustedUnitPrice);
                        row.createCell(offset + 9).setCellValue(sr.stabilitySubsidy);
                    }
                }
            }
        }
    }

    private void writeRewardSheet(Workbook wb, String title, List<Map<String, Object>> data) {
        Sheet sheet = wb.createSheet(title);
        String[] cols = {"姓名", "入职补贴（5元/人）", "稳岗补贴_驻场",
                "入职补贴（1元/人）", "稳岗补贴_经理", "合计", "备注"};
        Row header = sheet.createRow(0);
        for (int c = 0; c < cols.length; c++) header.createCell(c).setCellValue(cols[c]);
        for (int r = 0; r < data.size(); r++) {
            Row row = sheet.createRow(r + 1);
            Map<String, Object> map = data.get(r);
            row.createCell(0).setCellValue((String) map.getOrDefault("姓名", ""));
            row.createCell(1).setCellValue(((Number) map.getOrDefault("入职补贴（5元/人）", 0.0)).doubleValue());
            row.createCell(2).setCellValue(((Number) map.getOrDefault("稳岗补贴_驻场", 0.0)).doubleValue());
            row.createCell(3).setCellValue(((Number) map.getOrDefault("入职补贴（1元/人）", 0.0)).doubleValue());
            row.createCell(4).setCellValue(((Number) map.getOrDefault("稳岗补贴_经理", 0.0)).doubleValue());
            row.createCell(5).setCellValue(((Number) map.getOrDefault("合计", 0.0)).doubleValue());
            row.createCell(6).setCellValue((String) map.getOrDefault("备注", ""));
        }
    }

    private void writeBaseRateSheet(Workbook wb, String title, List<Map<String, Object>> data) {
        Sheet sheet = wb.createSheet(title);
        for (int r = 0; r < data.size(); r++) {
            Row row = sheet.createRow(r);
            Map<String, Object> map = data.get(r);
            for (int c = 1; c <= 13; c++) {
                Object val = map.getOrDefault("col" + c, "");
                Cell cell = row.createCell(c - 1);
                if (val instanceof Number) cell.setCellValue(((Number) val).doubleValue());
                else cell.setCellValue(String.valueOf(val));
            }
        }
    }

    private double roundHalfUp(double value, int decimals) {
        BigDecimal bd = BigDecimal.valueOf(value);
        return bd.setScale(decimals, RoundingMode.HALF_UP).doubleValue();
    }
}
