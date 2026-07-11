package com.example.project.service.finance.cxm;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 核对处理器 - 对应 Python: hs_check_processor.py
 * 功能：核算汇总Tab2"核对"的核心处理器
 *
 * 输出格式（3个Sheet）：
 *   Sheet1 "发放记录汇总" — 透视表: [供应商名称 | 各公司金额列 | 总计]
 *   Sheet2 "对账单汇总"   — 透视表: [供应商名称 | 各公司金额列 | 总计]  （格式与Sheet1一致，公司名已归一化）
 *   Sheet3 "核对结果"     — 对比表: [供应商名称 | 公司 | 发放记录金额 | 对账单金额 | 差额 | 状态]
 */
@Service
public class CheckProcessor {

    private static final Logger log = LoggerFactory.getLogger(CheckProcessor.class);

    // 空白字符正则
    private static final Pattern BLANK_PATTERN = Pattern.compile("[\\s\\u200b\\u200c\\u200d\\u00a0\\u3000]+");

    // 月份提取
    private static final Pattern MONTH_PATTERN = Pattern.compile("(?<!\\d)(\\d{1,2}月)(?!\\d)");

    // ==================== 参数检测接口 ====================

    /**
     * 检测发放记录表参数（列名、月份、类型）
     * 前端选择文件+工作表后调用，返回供用户选择的参数
     */
    public Map<String, Object> detectDisbursementParams(MultipartFile file, List<String> sheetNames) throws Exception {
        Map<String, Object> result = new HashMap<>();

        File inputFile = File.createTempFile("detect_disb_", ".xlsx");
        try {
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            Set<String> monthSet = new LinkedHashSet<>();
            List<Map<String, Object>> columnOptions = null;
            int defaultColH = -1;

            for (String sheetName : sheetNames) {
                SummaryData data = loadData(inputFile, sheetName);

                if (columnOptions == null) {
                    columnOptions = new ArrayList<>();
                    for (int i = 0; i < data.columns.size(); i++) {
                        Map<String, Object> colInfo = new HashMap<>();
                        colInfo.put("index", i);
                        colInfo.put("name", data.columns.get(i));
                        colInfo.put("colLetter", colIndexToLetter(i));
                        columnOptions.add(colInfo);
                    }
                    result.put("columns", columnOptions);

                    for (int i = 0; i < data.columns.size(); i++) {
                        if (data.columns.get(i).contains("金额") && defaultColH < 0) {
                            defaultColH = i;
                        }
                    }
                }

                // 合并所有选中工作表的月份
                int monthColIdx = findColumnIndex(data.columns, "备注1");
                if (monthColIdx < 0) monthColIdx = findColumnIndex(data.columns, "月份");

                if (monthColIdx >= 0) {
                    for (Object[] row : data.rows) {
                        if (monthColIdx < row.length && row[monthColIdx] != null) {
                            String val = row[monthColIdx].toString();
                            Matcher m = MONTH_PATTERN.matcher(val);
                            while (m.find()) monthSet.add(m.group(1));
                        }
                    }
                } else {
                    for (Object[] row : data.rows) {
                        for (Object cell : row) {
                            if (cell != null) {
                                Matcher m = MONTH_PATTERN.matcher(cell.toString());
                                while (m.find()) monthSet.add(m.group(1));
                            }
                        }
                    }
                }
            }

            result.put("defaultColH", defaultColH);

            List<String> months = new ArrayList<>(monthSet);
            months.sort((a, b) -> {
                int aNum = Integer.parseInt(a.replace("月", ""));
                int bNum = Integer.parseInt(b.replace("月", ""));
                return Integer.compare(aNum, bNum);
            });
            result.put("months", months);

            // 打款类型固定三种
            Set<String> existingTypes = new LinkedHashSet<>();
            for (String sheetName : sheetNames) {
                SummaryData data = loadData(inputFile, sheetName);
                int remark2Idx = findColumnIndex(data.columns, "备注2");
                if (remark2Idx >= 0) {
                    for (Object[] row : data.rows) {
                        if (remark2Idx < row.length && row[remark2Idx] != null) {
                            String val = row[remark2Idx].toString().trim();
                            if (val.contains("刘先锋现金卡")) existingTypes.add("刘先锋现金卡");
                            else if (val.endsWith("对公打款")) existingTypes.add("对公打款");
                            else if (val.endsWith("代发打款")) existingTypes.add("代发打款");
                        }
                    }
                }
            }
            List<String> types = new ArrayList<>();
            String[] typeOrder = {"刘先锋现金卡", "对公打款", "代发打款"};
            for (String t : typeOrder) {
                if (existingTypes.contains(t)) types.add(t);
            }
            result.put("types", types);

            result.put("success", true);
            log.info("检测参数完成（{}个工作表）：{} 列, {} 个月份, {} 个类型",
                    sheetNames.size(), columnOptions != null ? columnOptions.size() : 0,
                    months.size(), types.size());

        } finally {
            if (inputFile.exists()) inputFile.delete();
        }

        return result;
    }

    // ==================== 发放记录表汇总 ====================

    /**
     * 处理发放记录表汇总 → 输出透视表: [供应商名称 | 各公司列 | 总计]
     * 对应 Python: CheckProcessor.process_disbursement()
     */
    public String processDisbursement(MultipartFile file,
                                     String sheetName,
                                     int colH,
                                     String selectedMonths,
                                     String selectedTypes,
                                     String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("发放记录表不能为空");

        File inputFile = File.createTempFile("disbursement_", ".xlsx");
        try {
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            SummaryData data = loadData(inputFile, sheetName);

            List<String> monthsList = selectedMonths != null && !selectedMonths.isEmpty()
                    ? Arrays.asList(selectedMonths.split(",")) : null;
            List<String> typesList = selectedTypes != null && !selectedTypes.isEmpty()
                    ? Arrays.asList(selectedTypes.split(",")) : null;

            // 处理数据为透视表格式
            PivotResult result = processDisbursementPivot(data, colH, monthsList, typesList);

            // 写出透视表
            File outputFile = new File(outputDir, "发放记录汇总表.xlsx");
            writePivotTable(outputFile, result, "发放记录汇总", false); // false = 不含"公司"列

            log.info("发放记录表透视表完成：{} 行 x {} 列（含供应商名和总计）",
                    result.data.size(), result.columnHeaders.size());
            return outputFile.getAbsolutePath();

        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    // ==================== 对账单汇总 ====================

    /**
     * 处理对账单汇总 → 输出透视表: [供应商名称 | 各公司列 | 总计]
     * 注意：与发放记录汇总格式一致，不含"公司"列。公司名需归一化处理。
     * 对应 Python: CheckProcessor.process_statement()
     */
    public String processStatement(MultipartFile file,
                                  String sheetName,
                                  String selectedMonths,
                                  String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("对账单不能为空");

        File inputFile = File.createTempFile("statement_", ".xlsx");
        try {
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            SummaryData data = loadStatementData(inputFile, sheetName);

            List<String> monthsList = selectedMonths != null && !selectedMonths.isEmpty()
                    ? Arrays.asList(selectedMonths.split(",")) : null;

            // 处理数据为透视表格式（含公司列）
            PivotResult result = processStatementPivot(data, monthsList);

            // 写出透视表（与发放记录汇总格式一致，不含"公司"列）
            File outputFile = new File(outputDir, "对账单汇总表.xlsx");
            writePivotTable(outputFile, result, "对账单汇总", false); // false = 不含"公司"列

            log.info("对账单透视表完成：{} 行 x {} 列",
                    result.data.size(), result.columnHeaders.size());
            return outputFile.getAbsolutePath();

        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    // ==================== 核对对比 ====================

    /**
     * 执行核对对比 → 输出3个Sheet的Excel
     */
    public String reconcile(String disbursementFile,
                          String statementFile,
                          String outputDir) throws Exception {
        if (disbursementFile == null || !(new File(disbursementFile).exists())) {
            throw new IllegalArgumentException("发放记录汇总表不存在");
        }
        if (statementFile == null || !(new File(statementFile).exists())) {
            throw new IllegalArgumentException("对账单汇总表不存在");
        }

        File outputFile = new File(outputDir, "核对结果.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {

            // ---- Sheet1: 发放记录汇总（原样复制）----
            copySheetToFile(wb, disbursementFile, "发放记录汇总");

            // ---- Sheet2: 对账单汇总（原样复制）----
            copySheetToFile(wb, statementFile, "对账单汇总");

            // ---- Sheet3: 核对结果（逐格对比两个透视表）----
            ReconcileResult result = performReconcile(disbursementFile, statementFile);
            writeReconcileResult(wb, result);

            // 应用状态着色
            applyStatusColor(wb, "核对结果");

            log.info("核对完成：已核对={}, 待核对={}, 未匹配={}",
                    result.matched.size(), result.pending.size(), result.unmatched.size());

        } catch (Exception e) {
            log.error("核对失败", e);
            throw e;
        }

        return outputFile.getAbsolutePath();
    }

    // ========== 数据模型 ==========

    /** 通用数据结构 */
    static class SummaryData {
        String sheetName;
        List<String> columns; // 表头
        List<Object[]> rows;  // 数据行
    }

    /** 透视表结果 */
    static class PivotResult {
        List<String> columnHeaders;  // 所有列头: [供应商名称, 公司?, 台州分公司, ..., 总计]
        List<Map<String, Object>> data; // 每行是一个Map，key=列名, value=金额或null
    }

    /** 核对结果 */
    static class ReconcileResult {
        List<Map<String, Object>> matched = new ArrayList<>();   // 已核对
        List<Map<String, Object>> pending = new ArrayList<>();   // 待核对
        List<Map<String, Object>> unmatched = new ArrayList<>(); // 未匹配
    }

    // ========== 数据读取 ==========

    private SummaryData loadData(File file, String sheetName) throws Exception {
        SummaryData data = new SummaryData();
        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = findSheet(wb, sheetName);
            if (sheet == null) sheet = wb.getSheetAt(0);
            data.sheetName = sheet.getSheetName();

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new Exception("没有表头行");

            data.columns = new ArrayList<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                String val = getCellStringValue(headerRow.getCell(c));
                data.columns.add(BLANK_PATTERN.matcher(val.trim()).replaceAll(""));
            }

            data.rows = new ArrayList<>();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Object[] vals = new Object[data.columns.size()];
                for (int c = 0; c < data.columns.size(); c++) {
                    vals[c] = getCellValue(row.getCell(c));
                }
                data.rows.add(vals);
            }
        }
        return data;
    }

    /**
     * 读取对账单数据（可能结构与发放记录不同，但尝试用相同方式加载）
     */
    private SummaryData loadStatementData(File file, String sheetName) throws Exception {
        return loadData(file, sheetName);
    }

    private static Sheet findSheet(Workbook wb, String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) return wb.getSheetAt(0);
        String cleanName = BLANK_PATTERN.matcher(sheetName.trim()).replaceAll("");

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (BLANK_PATTERN.matcher(name.trim()).replaceAll("").equals(cleanName)) {
                return wb.getSheetAt(i);
            }
        }
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (name.toLowerCase().contains(sheetName.toLowerCase()) ||
                sheetName.toLowerCase().contains(name.toLowerCase())) {
                return wb.getSheetAt(i);
            }
        }
        return null;
    }

    // ========== 发放记录表处理（透视表）==========

    /**
     * 处理发放记录表数据为透视表格式
     * 输出列: [供应商名称, 台州分公司, 台州锦途, ..., 总计]
     *
     * 参考Python: SummaryProcessor.process_data() 的透视逻辑
     */
    private PivotResult processDisbursementPivot(SummaryData data,
                                                  int colH,
                                                  List<String> selectedMonths,
                                                  List<String> selectedTypes) {
        int idxSupplier = findColumnIndex(data.columns, "供应商名称");
        int idxRemark1 = findColumnIndex(data.columns, "备注1");
        int idxRemark2 = findColumnIndex(data.columns, "备注2");

        // ---- 第一步：筛选 + 提取公司缩写 + 收集所有唯一公司名 ----
        Set<String> companySet = new LinkedHashSet<>();

        // 中间数据：(原始行, 公司缩写)
        List<ProcessedRow> processedRows = new ArrayList<>();

        for (Object[] row : data.rows) {
            // 月份筛选
            if (selectedMonths != null && !selectedMonths.isEmpty()) {
                boolean monthMatch = false;
                if (idxRemark1 >= 0 && idxRemark1 < row.length && row[idxRemark1] != null) {
                    String remark1 = row[idxRemark1].toString();
                    for (String m : selectedMonths) {
                        if (remark1.contains(m)) { monthMatch = true; break; }
                    }
                }
                if (!monthMatch) continue;
            }

            // 类型筛选
            if (selectedTypes != null && !selectedTypes.isEmpty()) {
                boolean typeMatch = false;
                if (idxRemark2 >= 0 && idxRemark2 < row.length && row[idxRemark2] != null) {
                    String remark2 = row[idxRemark2].toString();
                    for (String t : selectedTypes) {
                        if (t.equals("刘先锋现金卡") && remark2.contains(t)) { typeMatch = true; break; }
                        if ((t.equals("对公打款") || t.equals("代发打款")) && remark2.endsWith(t)) { typeMatch = true; break; }
                    }
                }
                if (!typeMatch) continue;
            }

            // 获取供应商名称
            String supplier = idxSupplier >= 0 && idxSupplier < row.length && row[idxSupplier] != null
                    ? normalizeSupplierName(row[idxSupplier].toString()) : "";
            if (supplier.isEmpty()) continue;

            // 获取公司缩写（从备注2提取）
            String companyAbbrev = "";
            if (idxRemark2 >= 0 && idxRemark2 < row.length && row[idxRemark2] != null) {
                companyAbbrev = abbreviateCompanyName(row[idxRemark2].toString());
            }
            // 如果备注2为空白，归到"（空白）"
            if (companyAbbrev.isEmpty()) companyAbbrev = "（空白）";

            companySet.add(companyAbbrev);

            // 获取金额
            double amount = colH >= 0 && colH < row.length ? getNumericValue(row[colH]) : 0;

            ProcessedRow pr = new ProcessedRow();
            pr.supplier = supplier;
            pr.company = companyAbbrev;
            pr.amount = amount;
            pr.remark2Raw = idxRemark2 >= 0 && idxRemark2 < row.length && row[idxRemark2] != null
                    ? row[idxRemark2].toString() : "";
            processedRows.add(pr);
        }

        if (processedRows.isEmpty()) {
            throw new RuntimeException("没有符合条件的数据");
        }

        // "（空白）"排到最后
        List<String> companies = new ArrayList<>(companySet);
        if (companies.remove("（空白）")) companies.add("（空白）");

        // ---- 第二步：按供应商分组汇总到各公司列 ----
        // key = 供应商名称, value = {公司缩写 -> 金额合计}
        Map<String, Map<String, Double>> summaryMap = new LinkedHashMap<>();

        for (ProcessedRow pr : processedRows) {
            Map<String, Double> companyAmounts = summaryMap.computeIfAbsent(pr.supplier, k -> new HashMap<>());
            companyAmounts.merge(pr.company, pr.amount, Double::sum);
        }

        // ---- 第三步：构建输出透视表 ----
        List<Map<String, Object>> output = new ArrayList<>();
        for (Map.Entry<String, Map<String, Double>> entry : summaryMap.entrySet()) {
            Map<String, Object> row = new LinkedHashMap<>();
            row.put("供应商名称", entry.getKey());

            double total = 0;
            for (String company : companies) {
                Double amt = entry.getValue().get(company);
                double val = amt != null ? amt : 0;
                row.put(company, val > 0 ? val : null); // 0显示为空白
                total += val;
            }
            row.put("总计", total > 0 ? total : null);

            output.add(row);
        }

        // 排序
        output.sort((a, b) -> String.valueOf(a.get("供应商名称"))
                .compareTo(String.valueOf(b.get("供应商名称"))));

        // 组装列头
        PivotResult result = new PivotResult();
        result.columnHeaders = new ArrayList<>();
        result.columnHeaders.add("供应商名称");
        result.columnHeaders.addAll(companies);
        result.columnHeaders.add("总计");
        result.data = output;

        return result;
    }

    // ========== 对账单处理（透视表）==========

    /**
     * 处理对账单数据为透视表格式
     * 输出列: [供应商名称, 台州分公司, 台州锦途, ..., 总计]  （与发放记录汇总格式一致，不含"公司"列）
     *
     * 公司名归一化规则（对应Python核对Tab逻辑）：
     *   - "温州锦途外包" → "温州锦途"  （去"外包"后缀）
     *   - "合静"         → "温州合静"  （加"温州"前缀）
     */
    private PivotResult processStatementPivot(SummaryData data, List<String> selectedMonths) {
        int idxSupplier = findColumnIndex(data.columns, "供应商名称");
        int idxCompany = findColumnIndex(data.columns, "公司");
        int idxAmount = findColumnIndex(data.columns, "金额");
        int idxMonth = findColumnIndex(data.columns, "核算年月");
        if (idxMonth < 0) idxMonth = findColumnIndex(data.columns, "月份");

        // ---- 筛选 + 收集数据（公司名归一化）----
        Set<String> companySet = new LinkedHashSet<>();

        // (supplier, normalizedCompanyName, amount)
        List<StatementProcessedRow> processedRows = new ArrayList<>();

        for (Object[] row : data.rows) {
            // 月份筛选
            if (selectedMonths != null && !selectedMonths.isEmpty()) {
                boolean monthMatch = false;
                if (idxMonth >= 0 && idxMonth < row.length && row[idxMonth] != null) {
                    String monthVal = row[idxMonth].toString();
                    for (String m : selectedMonths) {
                        if (monthVal.contains(m)) { monthMatch = true; break; }
                    }
                }
                if (!monthMatch) continue;
            }

            // 供应商名称（归一化）
            String supplier = idxSupplier >= 0 && idxSupplier < row.length && row[idxSupplier] != null
                    ? normalizeSupplierName(row[idxSupplier].toString()) : "";
            if (supplier.isEmpty()) continue;

            // 公司名（从"公司"列获取，然后归一化处理）
            String companyForCol = "";

            if (idxCompany >= 0 && idxCompany < row.length && row[idxCompany] != null) {
                String rawCompany = row[idxCompany].toString().trim();
                // 应用公司名归一化规则
                companyForCol = normalizeStatementCompanyName(rawCompany);
            }

            // 如果没有"公司"列，尝试从备注2提取
            if (companyForCol.isEmpty()) {
                int idxRemark2 = findColumnIndex(data.columns, "备注2");
                if (idxRemark2 >= 0 && idxRemark2 < row.length && row[idxRemark2] != null) {
                    companyForCol = abbreviateCompanyName(row[idxRemark2].toString());
                }
            }
            if (companyForCol.isEmpty()) companyForCol = "（空白）";

            companySet.add(companyForCol);

            // 金额
            double amount = idxAmount >= 0 && idxAmount < row.length ? getNumericValue(row[idxAmount]) : 0;

            StatementProcessedRow spr = new StatementProcessedRow();
            spr.supplier = supplier;
            spr.companyForCol = companyForCol;
            spr.rawCompany = companyForCol; // 归一化后的名称
            spr.amount = amount;
            processedRows.add(spr);
        }

        if (processedRows.isEmpty()) {
            throw new RuntimeException("对账单中没有符合条件的数据");
        }

        List<String> companies = new ArrayList<>(companySet);
        if (companies.remove("（空白）")) companies.add("（空白）");

        // ---- 分组汇总（不含"公司"列，与发放记录汇总格式一致）----
        Map<String, Map<String, Double>> amountMap = new LinkedHashMap<>();

        for (StatementProcessedRow spr : processedRows) {
            Map<String, Double> amts = amountMap.computeIfAbsent(spr.supplier, k -> new HashMap<>());
            amts.merge(spr.companyForCol, spr.amount, Double::sum);
        }

        // ---- 构建输出透视表（不含"公司"列）----
        List<Map<String, Object>> output = new ArrayList<>();
        for (Map.Entry<String, Map<String, Double>> entry : amountMap.entrySet()) {
            Map<String, Object> rowData = new LinkedHashMap<>();
            rowData.put("供应商名称", entry.getKey());

            // 各公司金额列（不含额外的"公司"列）
            double total = 0;
            for (String company : companies) {
                Double amt = entry.getValue().get(company);
                double val = amt != null ? amt : 0;
                rowData.put(company, val > 0 ? val : null);
                total += val;
            }
            rowData.put("总计", total > 0 ? total : null);

            output.add(rowData);
        }

        output.sort((a, b) -> String.valueOf(a.get("供应商名称"))
                .compareTo(String.valueOf(b.get("供应商名称"))));

        PivotResult result = new PivotResult();
        result.columnHeaders = new ArrayList<>();
        result.columnHeaders.add("供应商名称");
        // 注意：不再添加"公司"列，与发放记录汇总格式一致
        result.columnHeaders.addAll(companies);
        result.columnHeaders.add("总计");
        result.data = output;

        return result;
    }

    // ========== 核对逻辑 ==========

    /**
     * 执行核对：读取两张透视表，逐格对比 (供应商, 公司) 的金额
     *
     * 对应 Python reconcile() 的状态判定逻辑：
     *   已核对 — 两边都有值且 A == B（严格相等）
     *   待核对 — 两边都有值但 A != B
     *   未匹配 — 只有一边有值（A有B无 或 A无B有）
     */
    private ReconcileResult performReconcile(String disbursementFile, String statementFile) throws Exception {
        ReconcileResult result = new ReconcileResult();

        // 读取发放记录透视表 → Map<(supplier, company), amount>
        Map<CellKey, Double> disbMap = loadPivotAsCellMap(disbursementFile, false);

        // 读取对账单透视表 → Map<(supplier, company), amount>
        Map<CellKey, Double> stmtMap = loadPivotAsCellMap(statementFile, false);

        // 收集所有的 (supplier, company) 组合
        Set<CellKey> allKeys = new LinkedHashSet<>();
        allKeys.addAll(disbMap.keySet());
        allKeys.addAll(stmtMap.keySet());

        for (CellKey key : allKeys) {
            Double disbAmt = disbMap.get(key);
            Double stmtAmt = stmtMap.get(key);

            double aVal = disbAmt != null ? disbAmt : 0.0;
            double bVal = stmtAmt != null ? stmtAmt : 0.0;
            double diff = aVal - bVal;

            // 双方都是0（或null），不显示（与Python一致：continue）
            if (Math.abs(aVal) < 1e-10 && Math.abs(bVal) < 1e-10) {
                continue;
            }

            Map<String, Object> row = new LinkedHashMap<>();
            row.put("供应商名称", key.supplier);
            row.put("公司", key.company);

            // 金额显示：0显示为空（与Python一致）
            row.put("发放记录金额", Math.abs(aVal) > 1e-10 ? aVal : "");
            row.put("对账单金额", Math.abs(bVal) > 1e-10 ? bVal : "");

            // 差额计算（与Python一致）
            if (Math.abs(aVal) > 1e-10 && Math.abs(bVal) > 1e-10) {
                // 两边都有值 → 显示差额
                row.put("差额", diff);
            } else if (Math.abs(aVal) > 1e-10) {
                // 只有发放记录有
                row.put("差额", aVal);
            } else {
                // 只有对账单有
                row.put("差额", -bVal);
            }

            // 状态判定（完全对齐Python逻辑）
            if (Math.abs(aVal) > 1e-10 && Math.abs(bVal) > 1e-10) {
                // 两边都有数据
                if (Math.abs(aVal - bVal) < 1e-10) {
                    // A == B → 已核对
                    row.put("状态", "已核对");
                    result.matched.add(row);
                } else {
                    // A != B → 待核对
                    row.put("状态", "待核对");
                    result.pending.add(row);
                }
            } else if (Math.abs(aVal) > 1e-10) {
                // 只有发放记录有 → 未匹配
                row.put("状态", "未匹配");
                result.unmatched.add(row);
            } else {
                // 只有对账单有 → 未匹配
                row.put("状态", "未匹配");
                result.unmatched.add(row);
            }
        }

        return result;
    }

    /**
     * 透视表单元格坐标（用于核对时定位每个格子的值）
     */
    static class CellKey {
        String supplier;
        String company;

        CellKey(String supplier, String company) {
            this.supplier = supplier;
            this.company = company;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof CellKey)) return false;
            return supplier.equals(((CellKey) o).supplier) && company.equals(((CellKey) o).company);
        }

        @Override
        public int hashCode() {
            return Objects.hash(supplier, company);
        }
    }

    /**
     * 将透视表文件加载为 Map<CellKey, Amount>
     * 遍历每一行的每一个公司列，提取 (供应商, 公司) → 金额
     */
    private Map<CellKey, Double> loadPivotAsCellMap(String filePath, boolean hasExtraCompanyCol) throws Exception {
        Map<CellKey, Double> map = new LinkedHashMap<>();

        try (Workbook wb = WorkbookFactory.create(new File(filePath))) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return map;

            // 读取表头，找出哪些是公司列
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.toString().trim());
            }

            // 确定公司列的范围：跳过前1~2列（供应商名称 / 公司），跳过最后的"总计"
            int colStart = hasExtraCompanyCol ? 2 : 1; // 有"公司"列则从第3列开始
            int colEnd = headers.size() - 1;             // 排除最后一列"总计"

            // 遍历数据行
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String supplier = "";
                Cell supplierCell = row.getCell(0);
                if (supplierCell != null) supplier = supplierCell.toString().trim();

                // 遍历每个公司列
                for (int c = colStart; c < colEnd; c++) {
                    if (c >= headers.size()) continue;
                    String company = headers.get(c);
                    if (company.isEmpty()) continue;

                    Cell cell = row.getCell(c);
                    if (cell == null) continue;

                    double amount = 0;
                    switch (cell.getCellType()) {
                        case NUMERIC: amount = cell.getNumericCellValue(); break;
                        case FORMULA:
                            try { amount = cell.getNumericCellValue(); } catch (Exception ignored) {}
                            break;
                        default: break;
                    }

                    // 保留所有非零值（包括负数，与Python一致）
                    if (Math.abs(amount) > 1e-10) {
                        map.put(new CellKey(supplier, company), amount);
                    }
                }
            }
        }

        return map;
    }

    // ========== 写出方法 ==========

    /**
     * 写出透视表到Excel文件
     * @param includeCompanyCol 是否包含"公司"作为第2列（对账单=true, 发放记录=false）
     */
    private void writePivotTable(File outputFile, PivotResult pivotResult,
                                  String sheetTitle, boolean includeCompanyCol) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet(sheetTitle);

            if (pivotResult.data.isEmpty()) {
                try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
                return;
            }

            // 写表头
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = createHeaderStyle(wb);
            List<String> headers = pivotResult.columnHeaders;
            for (int c = 0; c < headers.size(); c++) {
                Cell cell = headerRow.createCell(c);
                cell.setCellValue(headers.get(c));
                cell.setCellStyle(headerStyle);
            }

            // 写数据
            CellStyle dataStyle = createDataStyle(wb);
            for (int r = 0; r < pivotResult.data.size(); r++) {
                Row row = sheet.createRow(r + 1);
                Map<String, Object> rowData = pivotResult.data.get(r);
                for (int c = 0; c < headers.size(); c++) {
                    String colName = headers.get(c);
                    Object val = rowData.get(colName);
                    Cell cell = row.createCell(c);
                    if (val instanceof Number) {
                        cell.setCellValue(((Number) val).doubleValue());
                    } else if (val != null) {
                        cell.setCellValue(val.toString());
                    }
                    cell.setCellStyle(dataStyle);
                }
            }

            // 自动列宽
            for (int c = 0; c < headers.size(); c++) {
                sheet.autoSizeColumn(c);
                int current = sheet.getColumnWidth(c);
                sheet.setColumnWidth(c, Math.min(current + 1024, 50 * 256));
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
        }
    }

    /**
     * 将一个Excel文件的所有内容复制为目标Workbook的一个新Sheet
     */
    private void copySheetToFile(Workbook targetWb, String sourceFilePath, String sheetName) throws Exception {
        try (Workbook srcWb = WorkbookFactory.create(new File(sourceFilePath))) {
            Sheet srcSheet = srcWb.getSheetAt(0);
            Sheet destSheet = targetWb.createSheet(sheetName);

            // 复制表头和数据
            for (int r = 0; r <= srcSheet.getLastRowNum(); r++) {
                Row srcRow = srcSheet.getRow(r);
                if (srcRow == null) continue;

                Row destRow = destSheet.createRow(r);
                for (int c = 0; c < srcRow.getLastCellNum(); c++) {
                    Cell srcCell = srcRow.getCell(c);
                    if (srcCell == null) continue;

                    Cell destCell = destRow.createCell(c);
                    switch (srcCell.getCellType()) {
                        case NUMERIC:
                            destCell.setCellValue(srcCell.getNumericCellValue());
                            break;
                        case STRING:
                            destCell.setCellValue(srcCell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            destCell.setCellValue(srcCell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            try {
                                destCell.setCellValue(srcCell.getNumericCellValue());
                            } catch (Exception e) {
                                try {
                                    destCell.setCellValue(srcCell.getStringCellValue());
                                } catch (Exception ex) {
                                    destCell.setCellValue("");
                                }
                            }
                            break;
                        case BLANK:
                            break;
                        default:
                            destCell.setCellValue("");
                    }
                }
            }

            // 复制列宽
            for (int c = 0; c <= srcSheet.getLastRowNum(); c++) {
                destSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
            }
        }
    }

    /**
     * 写出核对结果Sheet: [供应商名称, 公司, 发放记录金额, 对账单金额, 差额, 状态]
     */
    private void writeReconcileResult(Workbook wb, ReconcileResult result) {
        Sheet sheet = wb.createSheet("核对结果");

        // 合并所有数据（保持顺序：已核对→待核对→未匹配）
        List<Map<String, Object>> allData = new ArrayList<>();
        allData.addAll(result.matched);
        allData.addAll(result.pending);
        allData.addAll(result.unmatched);

        String[] headers = {"供应商名称", "公司", "发放记录金额", "对账单金额", "差额", "状态"};

        // 表头
        Row headerRow = sheet.createRow(0);
        CellStyle headerStyle = createHeaderStyle(wb);
        for (int c = 0; c < headers.length; c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(headerStyle);
        }

        if (allData.isEmpty()) return;

        // 数据
        CellStyle dataStyle = createDataStyle(wb);
        for (int r = 0; r < allData.size(); r++) {
            Row row = sheet.createRow(r + 1);
            Map<String, Object> rowData = allData.get(r);
            for (int c = 0; c < headers.length; c++) {
                String colName = headers[c];
                Object val = rowData.get(colName);
                Cell cell = row.createCell(c);
                if (val instanceof Number) {
                    cell.setCellValue(((Number) val).doubleValue());
                } else if (val != null) {
                    cell.setCellValue(val.toString());
                }
                cell.setCellStyle(dataStyle);
            }
        }

        // 自动列宽
        for (int c = 0; c < headers.length; c++) {
            sheet.autoSizeColumn(c);
            int current = sheet.getColumnWidth(c);
            sheet.setColumnWidth(c, Math.min(current + 1024, 50 * 256));
        }
    }

    /**
     * 应用状态着色（仅对"核对结果"Sheet）
     * 已核对(绿) / 待核对(黄)
     */
    private void applyStatusColor(Workbook wb, String targetSheetName) {
        Sheet sheet = wb.getSheet(targetSheetName);
        if (sheet == null) return;

        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return;

        int statusColIdx = -1;
        for (Cell cell : headerRow) {
            if ("状态".equals(cell.toString().trim())) {
                statusColIdx = cell.getColumnIndex();
                break;
            }
        }
        if (statusColIdx < 0) return;

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Cell statusCell = row.getCell(statusColIdx);
            if (statusCell == null) continue;

            String status = statusCell.toString().trim();
            if (!"已核对".equals(status) && !"待核对".equals(status)) continue;

            CellStyle style = wb.createCellStyle();
            style.cloneStyleFrom(wb.getCellStyleAt((short) 0));

            if ("已核对".equals(status)) {
                ((XSSFCellStyle) style).setFillForegroundColor(
                        new XSSFColor(new byte[]{(byte)0x90, (byte)0xEE, (byte)0x90}, null));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else if ("待核对".equals(status)) {
                ((XSSFCellStyle) style).setFillForegroundColor(
                        new XSSFColor(new byte[]{(byte)0xFF, (byte)0xD7, (byte)0x00}, null));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            for (Cell cell : row) {
                cell.setCellStyle(style);
            }
        }
    }

    // ========== 工具方法 ==========

    private static String colIndexToLetter(int index) {
        StringBuilder sb = new StringBuilder();
        while (index >= 0) {
            sb.insert(0, (char) ('A' + (index % 26)));
            index = index / 26 - 1;
        }
        return sb.toString();
    }

    /**
     * 归一化供应商名称: 去"供应商-"前缀 + 去末尾数字 + 去"外包"后缀
     */
    private static String normalizeSupplierName(String name) {
        if (name == null || name.isEmpty()) return name;
        String normalized = name.replaceAll("^供应商-", "");
        normalized = normalized.replaceAll("\\d+$", "");
        normalized = normalized.replaceAll("外包$", "");
        return normalized.trim();
    }

    /**
     * 对账单公司名归一化（对应Python核对Tab逻辑）
     * 规则：
     *   1. "温州锦途外包" → "温州锦途"  （去"外包"后缀）
     *   2. "合静"         → "温州合静"  （加"温州"前缀）
     *   3. 其他公司名保持不变
     */
    private static String normalizeStatementCompanyName(String name) {
        if (name == null || name.isEmpty()) return "";

        String normalized = name.trim();

        // 规则1：去掉"外包"后缀
        if (normalized.endsWith("外包")) {
            normalized = normalized.substring(0, normalized.length() - "外包".length());
        }

        // 规则2："合静" → "温州合静"
        if ("合静".equals(normalized)) {
            normalized = "温州合静";
        }

        return normalized;
    }

    /**
     * 缩写公司名 + 归一化处理（对应 Python extract_company_and_type_from_remark2）
     * 步骤：
     *   1. 提取缩写："XXX对公打款" → "XXX"
     *   2. 归一化：去"外包"后缀 + "合静" → "温州合静"
     *
     * 示例：
     *   "温州锦途外包代发打款" → "温州锦途"
     *   "合静对公打款"         → "温州合静"
     *   "刘先锋现金卡支付"      → "刘先锋现金卡"
     */
    private static String abbreviateCompanyName(String remark2) {
        if (remark2 == null || remark2.isEmpty()) return "";

        String str = remark2.trim();

        // 特殊处理：刘先锋现金卡支付
        if (str.equals("刘先锋现金卡支付")) return "刘先锋现金卡";

        // 提取缩写：去掉后缀
        for (String suffix : new String[]{"对公打款", "代发打款"}) {
            if (str.endsWith(suffix)) str = str.substring(0, str.length() - suffix.length());
        }

        // 归一化1：去除末尾的"外包"
        if (str.endsWith("外包")) {
            str = str.substring(0, str.length() - "外包".length());
        }

        // 归一化2："合静" → "温州合静"
        if ("合静".equals(str)) {
            str = "温州合静";
        }

        return str.trim();
    }

    private static int findColumnIndex(List<String> columns, String name) {
        for (int i = 0; i < columns.size(); i++) {
            if (columns.get(i).contains(name)) return i;
        }
        return -1;
    }

    private static Object getCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return cell.getNumericCellValue();
            case BOOLEAN: return cell.getBooleanCellValue();
            case FORMULA:
                CellType rt = cell.getCachedFormulaResultType();
                if (rt == CellType.NUMERIC) return cell.getNumericCellValue();
                if (rt == CellType.STRING) return cell.getStringCellValue();
                return null;
            default: return null;
        }
    }

    private static String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                CellType rt = cell.getCachedFormulaResultType();
                if (rt == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
                if (rt == CellType.STRING) return cell.getStringCellValue();
                return "";
            default: return "";
        }
    }

    private static double getNumericValue(Object val) {
        if (val == null) return 0;
        if (val instanceof Number) return ((Number) val).doubleValue();
        try {
            return Double.parseDouble(val.toString());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static CellStyle createHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        org.apache.poi.ss.usermodel.Font font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        ((XSSFCellStyle)style).setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x44, (byte)0x72, (byte)0xC4}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createDataStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    // ========== 内部数据类 ==========

    /** 发放记录预处理行 */
    static class ProcessedRow {
        String supplier;
        String company;      // 公司缩写（用作透视表列名）
        double amount;       // 金额
        String remark2Raw;   // 原始备注2
    }

    /** 对账单预处理行 */
    static class StatementProcessedRow {
        String supplier;
        String companyForCol; // 用于透视表列名的公司标识
        String rawCompany;    // 原始公司名（用于"公司"列显示）
        double amount;
    }
}
