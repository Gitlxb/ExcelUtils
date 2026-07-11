package com.example.project.service.finance.cxm;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 对应 Python: hs_huizong_cxm.py + hs_huizong_cxm_main.py
 * 功能：核算汇总
 *
 * 核心处理流程：
 * 1. 从"备注2"列缩写公司名（刘先锋现金卡支付→刘先锋现金卡，XXX对公打款→XXX等）
 * 2. 按打款类型、月份、打款金额列筛选
 * 3. 按供应商+收款姓名分组汇总
 * 4. 生成汇总表（供应商名称 | 各公司金额列 | 总计 | 公司）
 *
 * 另含数据核对功能（DataChecker）
 */
@Service
public class HsHuizongCxmService {

    private static final org.slf4j.Logger log = org.slf4j.LoggerFactory.getLogger(HsHuizongCxmService.class);

    // 空白字符正则
    private static final Pattern BLANK_PATTERN = Pattern.compile("[\\s\\u200b\\u200c\\u200d\\u00a0\\u3000]+");
    // 月份提取
    private static final Pattern MONTH_PATTERN = Pattern.compile("(?<!\\d)(\\d{1,2}月)(?!\\d)");
    // 范围月份
    private static final Pattern RANGE_PATTERN = Pattern.compile("(?<!\\d)(\\d{1,2}-\\d{1,2}月)(?!\\d)");
    // 年份提取
    private static final Pattern YEAR_PATTERN = Pattern.compile("(\\d{4}年)");

    /**
     * 执行核算汇总
     * @param file               上传的 Excel 文件
     * @param selectedTypes      选择的打款类型（逗号分隔，如"刘先锋现金卡,对公打款"）
     * @param selectedMonths     选择的月份（逗号分隔，如"1月,2月"）
     * @param amountColIdx       打款金额列索引（0-based），-1表示自动使用第一个
     * @param outputDir          输出目录
     * @return 汇总结果文件路径
     */
    public String processSummary(MultipartFile file,
                                  String selectedTypes,
                                  String selectedMonths,
                                  int amountColIdx,
                                  String sheetName,
                                  String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        // amountColIdx == -1 时由前端传入选中的列索引，后端不再硬编码为0

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            // 1. 读取数据（支持指定工作表）
            SummaryData data = loadData(inputFile, sheetName);

            // 2. 解析用户选择
            List<String> typesList = selectedTypes != null && !selectedTypes.isEmpty()
                    ? Arrays.asList(selectedTypes.split(","))
                    : null;
            List<String> monthsList = selectedMonths != null && !selectedMonths.isEmpty()
                    ? Arrays.asList(selectedMonths.split(","))
                    : null;

            // 3. 处理数据
            List<Map<String, Object>> result = processData(data, typesList, monthsList, amountColIdx);

            // 4. 写出（传入选中月份以生成正确的工作表标题）
            File outputFile = new File(outputDir, "供应商拆分汇总.xlsx");
            writeSummary(outputFile, result, data.sheetName, selectedMonths);
            return outputFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    // ========== 数据模型 ==========

    static class SummaryData {
        String sheetName;
        List<String> columns; // 表头
        List<Object[]> rows;  // 数据行（每个元素为 Object[]，按 columns 索引）
    }

    // ========== 读取数据 ==========

    /** 模糊匹配工作表名称（去除隐藏字符后精确匹配，或忽略大小写包含匹配） */
    private static Sheet findSheet(Workbook wb, String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) return wb.getSheetAt(0);
        String cleanName = BLANK_PATTERN.matcher(sheetName.trim()).replaceAll("");
        // 第一遍：精确匹配（去除隐藏字符后）
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (BLANK_PATTERN.matcher(name.trim()).replaceAll("").equals(cleanName)) {
                return wb.getSheetAt(i);
            }
        }
        // 第二遍：忽略大小写包含匹配
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (name.toLowerCase().contains(sheetName.toLowerCase()) ||
                sheetName.toLowerCase().contains(name.toLowerCase())) {
                return wb.getSheetAt(i);
            }
        }
        return null;
    }

    private SummaryData loadData(File file, String sheetName) throws Exception {
        SummaryData data = new SummaryData();
        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = findSheet(wb, sheetName);
            if (sheet == null) {
                // 如果指定工作表不存在，使用第一个工作表
                sheet = wb.getSheetAt(0);
            }
            data.sheetName = sheet.getSheetName();

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new Exception("没有表头行");

            // 读取表头（去除隐藏字符，确保精确匹配）
            data.columns = new ArrayList<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                String val = ExcelUtils.getCellStringValue(headerRow.getCell(c));
                data.columns.add(BLANK_PATTERN.matcher(val.trim()).replaceAll(""));
            }

            // 必要列检查
            if (!data.columns.contains("备注2"))
                throw new Exception("未找到'备注2'列");
            if (!data.columns.contains("供应商名称"))
                throw new Exception("未找到'供应商名称'列");
            if (!data.columns.contains("收款姓名"))
                throw new Exception("未找到'收款姓名'列");

            // 读取数据行（含公式单元格处理）
            data.rows = new ArrayList<>();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Object[] vals = new Object[data.columns.size()];
                for (int c = 0; c < data.columns.size(); c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING: vals[c] = cell.getStringCellValue(); break;
                            case NUMERIC: vals[c] = cell.getNumericCellValue(); break;
                            case BOOLEAN: vals[c] = cell.getBooleanCellValue(); break;
                            case FORMULA:
                                // 公式单元格：根据缓存结果类型读取
                                CellType resultType = cell.getCachedFormulaResultType();
                                switch (resultType) {
                                    case STRING: vals[c] = cell.getStringCellValue(); break;
                                    case NUMERIC: vals[c] = cell.getNumericCellValue(); break;
                                    case BOOLEAN: vals[c] = cell.getBooleanCellValue(); break;
                                    default: vals[c] = null;
                                }
                                break;
                            default: vals[c] = null;
                        }
                    }
                }
                data.rows.add(vals);
            }
        }
        return data;
    }

    // ========== 列索引查找 ==========

    private int colIdx(List<String> columns, String name) {
        for (int i = 0; i < columns.size(); i++) {
            if (columns.get(i).equals(name)) return i;
        }
        return -1;
    }

    private List<Integer> colIdxContains(List<String> columns, String keyword) {
        List<Integer> result = new ArrayList<>();
        for (int i = 0; i < columns.size(); i++) {
            if (columns.get(i) != null && columns.get(i).contains(keyword)) result.add(i);
        }
        return result;
    }

    // ========== 公司名缩写 ==========

    private String abbreviateCompanyName(Object raw) {
        if (raw == null) return "（空白）";
        String s = raw.toString().trim();
        if (s.isEmpty()) return "（空白）";
        // 移除不可见空白
        s = BLANK_PATTERN.matcher(s).replaceAll("");
        if (s.isEmpty()) return "（空白）";

        // 规则1: "刘先锋现金卡支付" → "刘先锋现金卡"
        if ("刘先锋现金卡支付".equals(s)) return "刘先锋现金卡";

        // 规则2: "XXX对公打款" → "XXX" ; "XXX代发打款" → "XXX"
        if (s.endsWith("对公打款")) return s.replace("对公打款", "");
        if (s.endsWith("代发打款")) return s.replace("代发打款", "");

        return s;
    }

    private String extractPaymentType(Object remark2) {
        if (remark2 == null) return null;
        String s = remark2.toString().trim();
        if (s.contains("刘先锋现金卡")) return "刘先锋现金卡";
        if (s.endsWith("对公打款")) return "对公打款";
        if (s.endsWith("代发打款")) return "代发打款";
        return null;
    }

    // ========== 检测数据类型 ==========

    public Map<String, Object> detectColumns(MultipartFile file, String sheetName) throws Exception {
        Map<String, Object> result = new HashMap<>();

        File inputFile = File.createTempFile("input_", ".xlsx");
        try {
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            SummaryData data = loadData(inputFile, sheetName);

            // 检测打款类型
            List<String> types = new ArrayList<>();
            int idxRemark2 = colIdx(data.columns, "备注2");
            if (idxRemark2 >= 0) {
                Set<String> found = new LinkedHashSet<>();
                for (Object[] row : data.rows) {
                    if (idxRemark2 < row.length) {
                        String pt = extractPaymentType(row[idxRemark2]);
                        if (pt != null) found.add(pt);
                    }
                }
                types = new ArrayList<>(found);
                // 排序
                List<String> order = Arrays.asList("刘先锋现金卡", "对公打款", "代发打款");
                types.sort(Comparator.comparingInt(o -> order.indexOf(o) >= 0 ? order.indexOf(o) : 99));
            }

            // 检测月份
            List<String> months = new ArrayList<>();
            int idxRemark1 = colIdx(data.columns, "备注1");
            if (idxRemark1 >= 0) {
                Set<String> found = new LinkedHashSet<>();
                for (Object[] row : data.rows) {
                    if (idxRemark1 < row.length && row[idxRemark1] != null) {
                        String val = row[idxRemark1].toString();
                        Matcher m1 = MONTH_PATTERN.matcher(val);
                        while (m1.find()) found.add(m1.group(1));
                        Matcher m2 = RANGE_PATTERN.matcher(val);
                        while (m2.find()) found.add(m2.group(1));
                    }
                }
                // 排序：单月份在前，范围在后
                months = found.stream()
                        .sorted((a, b) -> {
                            boolean aRange = a.contains("-");
                            boolean bRange = b.contains("-");
                            if (aRange != bRange) return aRange ? 1 : -1;
                            int aNum = Integer.parseInt(a.replace("月", "").split("-")[0]);
                            int bNum = Integer.parseInt(b.replace("月", "").split("-")[0]);
                            return Integer.compare(aNum, bNum);
                        }).collect(Collectors.toList());
            }

            // 检测打款金额列
            List<Map<String, Object>> amountCols = new ArrayList<>();
            List<Integer> colIdxs = colIdxContains(data.columns, "打款金额");
            for (int i = 0; i < colIdxs.size(); i++) {
                Map<String, Object> colInfo = new HashMap<>();
                int idx = colIdxs.get(i);
                colInfo.put("index", idx);
                colInfo.put("name", data.columns.get(idx));
                colInfo.put("colLetter", colIndexToLetter(idx));
                amountCols.add(colInfo);
            }

            result.put("types", types);
            result.put("months", months);
            result.put("amountColumns", amountCols);
            result.put("columns", data.columns);
            result.put("success", true);

        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
        return result;
    }

    // ========== 核心处理 ==========

    private List<Map<String, Object>> processData(SummaryData data,
                                                   List<String> selectedTypes,
                                                   List<String> selectedMonths,
                                                   int amountColIdx) throws Exception {
        int idxRemark2 = colIdx(data.columns, "备注2");
        int idxRemark1 = colIdx(data.columns, "备注1");
        int idxSupplier = colIdx(data.columns, "供应商名称");
        int idxPayee = colIdx(data.columns, "收款姓名");

        // 1. 添加公司缩写列（按索引扩展）
        List<String> extendedCols = new ArrayList<>(data.columns);
        int idxAbbrev = extendedCols.size();
        extendedCols.add("_公司缩写");

        // 2. 逐行处理
        List<Object[]> processedRows = new ArrayList<>();
        for (Object[] row : data.rows) {
            boolean pass = true;

            // 类型筛选
            if (selectedTypes != null && !selectedTypes.isEmpty()) {
                String pt = extractPaymentType(idxRemark2 < row.length ? row[idxRemark2] : null);
                pass = pt != null && selectedTypes.contains(pt);
            }
            if (!pass) continue;

            // 月份筛选
            if (selectedMonths != null && !selectedMonths.isEmpty()) {
                String remark1 = idxRemark1 < row.length && row[idxRemark1] != null
                        ? row[idxRemark1].toString() : "";
                pass = selectedMonths.stream().anyMatch(m -> matchMonthInText(m, remark1));
            }
            if (!pass) continue;

            // 扩展行（加入缩写列）
            Object[] extRow = Arrays.copyOf(row, extendedCols.size());
            extRow[idxAbbrev] = abbreviateCompanyName(idxRemark2 < row.length ? row[idxRemark2] : null);
            processedRows.add(extRow);
        }

        if (processedRows.isEmpty()) {
            throw new Exception("没有符合筛选条件的数据");
        }

        // 3. 获取所有唯一的缩写公司名
        Set<String> companySet = new LinkedHashSet<>();
        for (Object[] row : processedRows) {
            String abbrev = (String) row[idxAbbrev];
            if (abbrev != null && !abbrev.isEmpty()) companySet.add(abbrev);
        }
        // "（空白）"放最后
        if (companySet.remove("（空白）")) companySet.add("（空白）");

        List<String> companies = new ArrayList<>(companySet);

        // 4. 按供应商+收款姓名分组汇总
        Map<String, SummaryRow> summaryMap = new LinkedHashMap<>();

        for (Object[] row : processedRows) {
            String supplier = row[idxSupplier] != null ? row[idxSupplier].toString().trim() : "";
            if (supplier.isEmpty()) continue;
            // 空白字符容错：移除不可见字符（对应 Python _clean_blank_chars）
            supplier = BLANK_PATTERN.matcher(supplier).replaceAll("");

            String payee = idxPayee < row.length && row[idxPayee] != null
                    ? row[idxPayee].toString().trim() : "";
            payee = BLANK_PATTERN.matcher(payee).replaceAll("");

            String key = payee.isEmpty() ? supplier : supplier + "|" + payee;
            // 使用传统 get+put，避开 lambda effectively final 问题
            SummaryRow sr = summaryMap.get(key);
            if (sr == null) {
                sr = new SummaryRow();
                sr.supplier = supplier;
                sr.payee = payee;
                sr.amounts = new HashMap<>();
                summaryMap.put(key, sr);
            }

            String abbrev = (String) row[idxAbbrev];
            // 空白备注2归到"（空白）"
            boolean isBlankRemark = row[idxRemark2] == null
                    || BLANK_PATTERN.matcher(row[idxRemark2].toString()).replaceAll("").isEmpty();

            if (isBlankRemark) {
                // 空白行：汇总到"（空白）"列
                double amt = getAmount(row, amountColIdx);
                sr.amounts.merge("（空白）", amt, Double::sum);
            } else {
                // 非空白行：汇总到对应公司列
                double amt = getAmount(row, amountColIdx);
                sr.amounts.merge(abbrev, amt, Double::sum);
            }
        }

        // 5. 构建输出
        List<Map<String, Object>> output = new ArrayList<>();
        for (Map.Entry<String, SummaryRow> entry : summaryMap.entrySet()) {
            SummaryRow sr = entry.getValue();
            Map<String, Object> row = new LinkedHashMap<>();
            row.put("供应商名称", sr.supplier);

            double total = 0;
            for (String company : companies) {
                double amt = sr.amounts.getOrDefault(company, 0.0);
                row.put(company, amt > 0 ? amt : null); // 0 不显示
                total += amt;
            }
            row.put("总计", total > 0 ? total : null);

            // 公司列：如果收款姓名含关键词则显示
            row.put("备注", null); // 空列分隔
            row.put("公司", extractCompanyFromPayee(sr.payee));

            output.add(row);
        }

        // 按供应商名称排序
        output.sort((a, b) -> String.valueOf(a.get("供应商名称")).compareTo(String.valueOf(b.get("供应商名称"))));

        return output;
    }

    // ========== 多Sheet批量处理 ==========
    // 对应 Python: hs_gui_actions.py::_process_data() 中多Sheet处理部分

    /**
     * 批量处理多个工作表，并合并结果
     * @param file               上传的 Excel 文件（含多个工作表）
     * @param sheetNames        选择的工作表名称列表
     * @param selectedTypes     选择的打款类型
     * @param selectedMonths   选择的月份
     * @param amountColIdx    打款金额列索引
     * @param outputDir        输出目录
     * @return 汇总结果文件路径
     */
    public String processSummaryBatch(MultipartFile file,
                                       List<String> sheetNames,
                                       String selectedTypes,
                                       String selectedMonths,
                                       int amountColIdx,
                                       String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (sheetNames == null || sheetNames.isEmpty()) throw new IllegalArgumentException("请选择至少一个工作表");

        File inputFile = File.createTempFile("batch_", ".xlsx");
        try {
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            // 逐个处理每个工作表
            List<List<Map<String, Object>>> allResults = new ArrayList<>();
            for (String sheetName : sheetNames) {
                try {
                    // 读取数据
                    SummaryData data = loadData(inputFile, sheetName);

                    // 解析用户选择
                    List<String> typesList = selectedTypes != null && !selectedTypes.isEmpty()
                            ? Arrays.asList(selectedTypes.split(","))
                            : null;
                    List<String> monthsList = selectedMonths != null && !selectedMonths.isEmpty()
                            ? Arrays.asList(selectedMonths.split(","))
                            : null;

                    // 处理数据
                    List<Map<String, Object>> result = processData(data, typesList, monthsList, amountColIdx);
                    allResults.add(result);

                    log.info("工作表 [{}] 处理完成，结果行数: {}", sheetName, result.size());
                } catch (Exception e) {
                    log.warn("工作表 [{}] 处理失败: {}", sheetName, e.getMessage());
                    // 继续处理其他工作表
                }
            }

            if (allResults.isEmpty()) {
                throw new Exception("所有工作表处理失败，请检查文件格式");
            }

            // 合并所有结果
            List<Map<String, Object>> mergedResult = mergeResults(allResults);

            // 写出（使用固定文件名）
            File outputFile = new File(outputDir, "供应商拆分汇总.xlsx");
            // 合并后的工作表标题：如果有选择月份，则使用第一个月份
            String sheetTitle = selectedMonths != null && !selectedMonths.isEmpty()
                    ? selectedMonths.split(",")[0]
                    : "供应商汇总";
            writeSummary(outputFile, mergedResult, sheetNames.get(0), selectedMonths);

            log.info("批量处理完成，共处理 {} 个工作表，合并后行数: {}", sheetNames.size(), mergedResult.size());
            return outputFile.getAbsolutePath();

        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    /**
     * 合并多个处理结果
     * 对应 Python: 多个DataFrame concat 后按 供应商名称+公司 分组求和
     */
    private List<Map<String, Object>> mergeResults(List<List<Map<String, Object>>> allResults) {
        // 收集所有列名
        Set<String> allColumns = new LinkedHashSet<>();
        for (List<Map<String, Object>> result : allResults) {
            for (Map<String, Object> row : result) {
                allColumns.addAll(row.keySet());
            }
        }

        // 按 供应商名称+公司 分组求和
        Map<String, Map<String, Object>> mergedMap = new LinkedHashMap<>();

        for (List<Map<String, Object>> result : allResults) {
            for (Map<String, Object> row : result) {
                String supplier = (String) row.get("供应商名称");
                String company = (String) row.get("公司");

                // 使用 供应商名称+公司 作为合并键
                String key = supplier + "|" + (company != null ? company : "");

                Map<String, Object> mergedRow = mergedMap.get(key);
                if (mergedRow == null) {
                    mergedRow = new LinkedHashMap<>();
                    mergedRow.put("供应商名称", supplier);
                    mergedRow.put("公司", company);
                    // 初始化所有金额列为 0
                    for (String col : allColumns) {
                        if (!"供应商名称".equals(col) && !"公司".equals(col)) {
                            mergedRow.put(col, 0.0);
                        }
                    }
                    mergedMap.put(key, mergedRow);
                }

                // 累加金额列
                for (String col : allColumns) {
                    if ("供应商名称".equals(col) || "公司".equals(col)) continue;

                    Object val = row.get(col);
                    if (val instanceof Number) {
                        double current = ((Number) mergedRow.get(col)).doubleValue();
                        double add = ((Number) val).doubleValue();
                        mergedRow.put(col, current + add);
                    }
                }
            }
        }

        // 转换为 List，并将 0 替换为 null（不显示）
        List<Map<String, Object>> output = new ArrayList<>();
        for (Map<String, Object> row : mergedMap.values()) {
            Map<String, Object> cleanedRow = new LinkedHashMap<>();
            for (Map.Entry<String, Object> entry : row.entrySet()) {
                Object val = entry.getValue();
                if (val instanceof Number) {
                    double d = ((Number) val).doubleValue();
                    cleanedRow.put(entry.getKey(), d > 0.001 ? d : null);
                } else {
                    cleanedRow.put(entry.getKey(), val);
                }
            }
            output.add(cleanedRow);
        }

        // 按供应商名称排序
        output.sort((a, b) -> String.valueOf(a.get("供应商名称")).compareTo(String.valueOf(b.get("供应商名称"))));

        return output;
    }

    static class SummaryRow {
        String supplier;
        String payee;
        Map<String, Double> amounts;
    }

    private double getAmount(Object[] row, int idx) {
        if (idx < row.length && row[idx] != null) {
            try {
                return Double.parseDouble(row[idx].toString());
            } catch (NumberFormatException ignored) {}
        }
        return 0;
    }

    private boolean matchMonthInText(String month, String text) {
        if (text == null || text.isEmpty()) return false;
        Pattern p = Pattern.compile("(?<!\\d)" + Pattern.quote(month) + "(?!\\d)");
        return p.matcher(text).find();
    }

    private static final List<String> COMPANY_KEYWORDS = Arrays.asList(
            "刘先锋现金卡支付", "个体工商户", "个体户", "有限公司", "分公司", "服务部");

    private String extractCompanyFromPayee(String payee) {
        if (payee == null || payee.isEmpty()) return "";
        for (String kw : COMPANY_KEYWORDS) {
            if (payee.contains(kw)) return payee;
        }
        return "";
    }

    // ========== 写出 ==========

    private void writeSummary(File outputFile, List<Map<String, Object>> data,
                                String sheetName, String selectedMonths) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet(extractSheetTitle(sheetName, selectedMonths));

            if (data.isEmpty()) {
                try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
                return;
            }

            // 获取所有列，按固定顺序排列：
            // 供应商名称 → [各公司金额列] → 总计 → 备注 → 公司
            Set<String> colSet = new LinkedHashSet<>();
            for (Map<String, Object> row : data) colSet.addAll(row.keySet());
            List<String> columns = new ArrayList<>(colSet);
            // 移除固定位置的特殊列
            columns.remove("供应商名称");
            columns.remove("公司");
            columns.remove("总计");
            columns.remove("备注");
            columns.remove(""); // 去掉空串列
            // 按正确顺序组装
            List<String> orderedColumns = new ArrayList<>();
            orderedColumns.add("供应商名称");
            orderedColumns.addAll(columns); // 动态公司金额列（刘先锋现金卡、（空白）等）
            orderedColumns.add("总计");
            orderedColumns.add("备注");
            orderedColumns.add("公司");
            columns = orderedColumns;

            // 写表头
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = ExcelUtils.createHeaderStyle(wb);
            for (int c = 0; c < columns.size(); c++) {
                Cell cell = headerRow.createCell(c);
                cell.setCellValue(columns.get(c));
                cell.setCellStyle(headerStyle);
            }

            // 写数据
            for (int r = 0; r < data.size(); r++) {
                Row row = sheet.createRow(r + 1);
                Map<String, Object> map = data.get(r);
                for (int c = 0; c < columns.size(); c++) {
                    Object val = map.get(columns.get(c));
                    Cell cell = row.createCell(c);
                    if (val instanceof Number) {
                        cell.setCellValue(((Number) val).doubleValue());
                    } else if (val != null) {
                        cell.setCellValue(val.toString());
                    }
                    // 居中
                    cell.setCellStyle(getDataCellStyle(wb));
                }
            }

            // 应用统一样式（表头样式、数据边框、自动列宽）
            applySummaryStyles(wb, sheet);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
        }
    }

    private CellStyle getDataCellStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    // ========== Excel 样式（对应 hs_utils.py -> apply_excel_styles）==========

    /**
     * 对汇总输出 Sheet 应用统一样式
     * 对应 Python: apply_excel_styles()
     */
    public static void applySummaryStyles(Workbook wb, Sheet sheet) {
        CellStyle headerStyle = createSummaryHeaderStyle(wb);
        CellStyle dataStyle = createSummaryDataStyle(wb);

        // 表头样式
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                cell.setCellStyle(headerStyle);
            }
        }

        // 数据行样式 + 自动列宽
        int lastCol = 0;
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (Cell cell : row) {
                cell.setCellStyle(dataStyle);
                lastCol = Math.max(lastCol, cell.getColumnIndex());
            }
        }

        // 自动列宽（最大 50 字符）
        for (int c = 0; c <= lastCol; c++) {
            sheet.autoSizeColumn(c);
            int current = sheet.getColumnWidth(c);
            sheet.setColumnWidth(c, Math.min(current + 1024, 50 * 256));
        }
    }

    private static CellStyle createSummaryHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        org.apache.poi.ss.usermodel.Font font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        // #4472c4 = RGB(68, 114, 196)
        ((XSSFCellStyle)style).setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x44, (byte)0x72, (byte)0xC4}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createSummaryDataStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    // ========== 追加到目标表（对应 hs_gui.py -> _append_to_target）==========

    /**
     * 将汇总数据和核对结果追加到目标 Excel 文件
     * 若目标文件已存在，则追加为新的工作表（同名时自动加 _1, _2...）
     * 对应 Python: save_combined_excel()
     *
     * @param summaryFile  汇总结果文件（已由 writeSummary 生成）
     * @param checkFile    核对结果文件（已由 HsHuizongCxmChecker 生成）
     * @param targetPath   目标文件路径
     * @return 目标文件绝对路径
     */
    /**
     * 将汇总结果和核对结果追加到目标文件
     * 对应 Python: hs_gui_actions.py::_append_to_target()
     * 增加重复防护：目标文件已包含同类型Sheet时跳过并提示
     *
     * @return Map 包含 targetPath, appended(已追加列表), skipped(已跳过列表), message
     */
    public Map<String, Object> appendToTarget(String summaryFile, String checkFile, String targetPath) throws Exception {
        Map<String, Object> result = new LinkedHashMap<>();
        File target = new File(targetPath);
        boolean targetExists = target.exists();

        List<String> appended = new ArrayList<>();
        List<String> skipped = new ArrayList<>();

        try (Workbook wb = targetExists
                ? new XSSFWorkbook(new FileInputStream(target))
                : new XSSFWorkbook()) {

            // 收集已有的 Sheet 名称（用于重复检测）
            Set<String> existing = new HashSet<>();
            if (targetExists) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    existing.add(wb.getSheetName(i));
                }
            }

            // 检查是否已有同类 Sheet（前缀匹配）
            boolean hasSummary = hasSheetWithPrefix(existing, "供应商拆分汇总");
            boolean hasCheck = hasSheetWithPrefix(existing, "核对结果");

            // 复制汇总表（如果不存在同类）
            if (summaryFile != null && new File(summaryFile).exists()) {
                if (hasSummary) {
                    skipped.add("供应商拆分汇总");
                    log.info("目标文件已包含供应商拆分汇总类工作表，跳过");
                } else {
                    String summarySheetName = getUniqueSheetName(existing, "供应商拆分汇总");
                    copySheetToFile(wb, summaryFile, summarySheetName);
                    existing.add(summarySheetName);
                    appended.add("供应商拆分汇总");
                }
            }

            // 复制核对结果表（如果不存在同类）
            if (checkFile != null && new File(checkFile).exists()) {
                if (hasCheck) {
                    skipped.add("核对结果");
                    log.info("目标文件已包含核对结果类工作表，跳过");
                } else {
                    String checkSheetName = getUniqueSheetName(existing, "核对结果");
                    copySheetToFile(wb, checkFile, checkSheetName);
                    existing.add(checkSheetName);
                    appended.add("核对结果");
                }
            }

            // 保存
            try (FileOutputStream fos = new FileOutputStream(target)) {
                wb.write(fos);
            }
        }

        // 注意：不删除源文件（供应商拆分汇总.xlsx / 核对结果.xlsx）
        // 用户追加后仍需单独下载这两个文件，删除会导致下载失效
        // 临时文件由 Controller 层负责清理

        // 组装返回信息
        result.put("targetPath", target.getAbsolutePath());
        result.put("appended", appended);
        result.put("skipped", skipped);

        if (!appended.isEmpty() && !skipped.isEmpty()) {
            result.put("message", "部分数据已追加（" + String.join("、", appended) +
                    "），部分已存在跳过（" + String.join("、", skipped) + "）");
        } else if (!appended.isEmpty()) {
            result.put("message", "追加完成：" + String.join("、", appended));
        } else if (!skipped.isEmpty()) {
            result.put("message", "所有可用数据均已存在于目标表中，无需重复追加（" +
                    String.join("、", skipped) + "）");
        } else {
            result.put("message", "没有可追加的数据");
        }

        return result;
    }

    /** 检查是否有以指定前缀开头的 Sheet */
    private static boolean hasSheetWithPrefix(Set<String> sheetNames, String prefix) {
        for (String name : sheetNames) {
            if (name.startsWith(prefix) || name.equals(prefix)) return true;
        }
        return false;
    }

    /** 生成不重复的工作表名 */
    private static String getUniqueSheetName(Set<String> existing, String base) {
        if (!existing.contains(base)) return base;
        int i = 1;
        while (existing.contains(base + "_" + i)) i++;
        return base + "_" + i;
    }

    /** 将单个 Excel 文件复制为指定名称的工作表到目标 Workbook */
    private static void copySheetToFile(Workbook targetWb, String sourceFilePath, String targetSheetName) throws Exception {
        try (Workbook srcWb = new XSSFWorkbook(new FileInputStream(sourceFilePath))) {
            Sheet srcSheet = srcWb.getSheetAt(0);
            Sheet targetSheet = targetWb.createSheet(targetSheetName);

            // 复制列宽
            for (int c = 0; c < srcSheet.getRow(0).getLastCellNum(); c++) {
                targetSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
            }

            // 复制行
            for (int r = 0; r <= srcSheet.getLastRowNum(); r++) {
                Row srcRow = srcSheet.getRow(r);
                if (srcRow == null) continue;
                Row targetRow = targetSheet.createRow(r);
                targetRow.setHeight(srcRow.getHeight());
                for (Cell srcCell : srcRow) {
                    Cell targetCell = targetRow.createCell(srcCell.getColumnIndex());
                    copyCellValue(srcCell, targetCell);
                    copyCellStyle(srcCell, targetCell, targetWb);
                }
            }
        }
    }

    private static void copyCellValue(Cell src, Cell target) {
        switch (src.getCellType()) {
            case STRING: target.setCellValue(src.getStringCellValue()); break;
            case NUMERIC: target.setCellValue(src.getNumericCellValue()); break;
            case BOOLEAN: target.setCellValue(src.getBooleanCellValue()); break;
            case FORMULA: target.setCellFormula(src.getCellFormula()); break;
            default: break;
        }
    }

    private static void copyCellStyle(Cell src, Cell target, Workbook targetWb) {
        CellStyle newStyle = targetWb.createCellStyle();
        CellStyle srcStyle = src.getCellStyle();
        newStyle.cloneStyleFrom(srcStyle);
        target.setCellStyle(newStyle);
    }

    /** 从工作表名称中提取标题（如"2026年1月"或"供应商汇总"）
     *  对应 Python: _extract_date_from_sheet_name()
     *  增强：若传入 selectedMonth，则拼接到年份后
     */
    private String extractSheetTitle(String sheetName, String selectedMonth) {
        if (sheetName == null || sheetName.isEmpty()) {
            return selectedMonth != null ? selectedMonth : "供应商汇总";
        }
        java.util.regex.Matcher m = YEAR_PATTERN.matcher(sheetName);
        if (m.find()) {
            String year = m.group(1); // e.g. "2026年"
            if (selectedMonth != null && !selectedMonth.isEmpty()) {
                // selectedMonth 如 "1月" 或 "1-2月"
                return year + selectedMonth;
            }
            return year;
        }
        // 无法提取年份，直接返回 selectedMonth 或默认名
        return selectedMonth != null && !selectedMonth.isEmpty() ? selectedMonth : "供应商汇总";
    }

    /** 将列索引转换为 Excel 列字母（如 0->A, 2->C, 7->H, 26->AA） */
    private static String colIndexToLetter(int col) {
        StringBuilder sb = new StringBuilder();
        while (col >= 0) {
            sb.insert(0, (char) ('A' + (col % 26)));
            col = col / 26 - 1;
        }
        return sb.toString();
    }

    // ========= 批量检测列信息（多工作表）=========
    /**
     * 批量检测多个工作表的列信息，合并结果
     * 对应 Python: hs_gui_actions.py::_process_data() 中的预扫描逻辑
     */
    public Map<String, Object> detectColumnsBatch(MultipartFile file, List<String> sheetNames) throws Exception {
        Map<String, Object> mergedResult = new HashMap<>();
        Set<String> allTypes = new LinkedHashSet<>();
        Set<String> allMonths = new LinkedHashSet<>();
        List<Map<String, Object>> allAmountCols = new ArrayList<>();

        File inputFile = File.createTempFile("batch_detect_", ".xlsx");
        try {
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(inputFile)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            for (String sheetName : sheetNames) {
                try {
                    SummaryData data = loadData(inputFile, sheetName);

                    // 检测打款类型
                    int idxRemark2 = colIdx(data.columns, "备注2");
                    if (idxRemark2 >= 0) {
                        for (Object[] row : data.rows) {
                            if (idxRemark2 < row.length) {
                                String pt = extractPaymentType(row[idxRemark2]);
                                if (pt != null) allTypes.add(pt);
                            }
                        }
                    }

                    // 检测月份
                    int idxRemark1 = colIdx(data.columns, "备注1");
                    if (idxRemark1 >= 0) {
                        for (Object[] row : data.rows) {
                            if (idxRemark1 < row.length && row[idxRemark1] != null) {
                                String val = row[idxRemark1].toString();
                                Matcher m1 = MONTH_PATTERN.matcher(val);
                                while (m1.find()) allMonths.add(m1.group(1));
                                Matcher m2 = RANGE_PATTERN.matcher(val);
                                while (m2.find()) allMonths.add(m2.group(1));
                            }
                        }
                    }

                    // 检测打款金额列
                    List<Integer> colIdxs = colIdxContains(data.columns, "打款金额");
                    for (int idx : colIdxs) {
                        boolean exists = false;
                        for (Map<String, Object> col : allAmountCols) {
                            if ((Integer) col.get("index") == idx) { exists = true; break; }
                        }
                        if (!exists) {
                            Map<String, Object> colInfo = new HashMap<>();
                            colInfo.put("index", idx);
                            colInfo.put("name", data.columns.get(idx));
                            colInfo.put("colLetter", colIndexToLetter(idx));
                            allAmountCols.add(colInfo);
                        }
                    }

                } catch (Exception e) {
                    log.warn("工作表 [{}] 检测失败: {}", sheetName, e.getMessage());
                }
            }
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }

        // 排序
        List<String> typesList = new ArrayList<>(allTypes);
        List<String> order = Arrays.asList("刘先锋现金卡", "对公打款", "代发打款");
        typesList.sort(Comparator.comparingInt(o -> order.indexOf(o) >= 0 ? order.indexOf(o) : 99));

        List<String> monthsList = new ArrayList<>(allMonths);
        monthsList.sort((a, b) -> {
            boolean aRange = a.contains("-");
            boolean bRange = b.contains("-");
            if (aRange != bRange) return aRange ? 1 : -1;
            int aNum = Integer.parseInt(a.replace("月", "").split("-")[0]);
            int bNum = Integer.parseInt(b.replace("月", "").split("-")[0]);
            return Integer.compare(aNum, bNum);
        });

        mergedResult.put("types", typesList);
        mergedResult.put("months", monthsList);
        mergedResult.put("amountColumns", allAmountCols);
        mergedResult.put("success", true);

        log.info("批量检测完成：类型={}, 月份={}, 金额列={}", typesList, monthsList, allAmountCols.size());
        return mergedResult;
    }

    // ==================== 数据核对（委托给 HsHuizongCxmChecker）====================

    /**
     * 核对数据：对比待核对表与基准表的金额
     * 实际逻辑由 HsHuizongCxmChecker 完成
     * 此方法保留为了兼容 Controller 签名
     */
    public String checkData(MultipartFile refFile, MultipartFile checkFile,
                             String checkSheet, String selectedMonth,
                             String selectedPersons, String outputDir) throws Exception {
        throw new UnsupportedOperationException("请使用 /summary/check 接口");
    }
}
