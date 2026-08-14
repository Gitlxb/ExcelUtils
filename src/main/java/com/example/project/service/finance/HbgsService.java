package com.example.project.service.finance;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 对应 Python: hbgs_core.py
 * 功能：合并个税统计表，按手续费规则计算银行手续费
 *
 * 核心流程（与 Python run_merge_collect 完全等价）：
 * 1. 加载手续费规则表（load_fee_rules）
 * 2. 扫描所有Excel文件（scan_excel_files）
 * 3. 逐个读取「个税统计」表 + 「员工信息-X月」表（_read_payroll_file）
 * 4. 按规则逐行计算手续费（compute_fee_v2）
 * 5. 输出汇总表，身份证列文本化（write_merge_output）
 */
@Service
public class HbgsService {

    private static final Logger log = LoggerFactory.getLogger(HbgsService.class);

    // ============================================================
    // 手续费规则分类：A~M 各类精确匹配串（与 VBA 原值逐字一致，含全角标点）
    // 对应 Python _RULE_SETS
    // ============================================================

    private static final Map<String, Set<String>> RULE_SETS = new LinkedHashMap<>();

    static {
        RULE_SETS.put("A", new LinkedHashSet<>(Arrays.asList(
                "只有代工费", "只有管理费，没有工资", "只有代工费，没有工资", "不扣手续费",
                "日结", "联动代发不收手续费", "不发工资，没有手续费", "不扣手续费，锦途发放",
                "不收手续费", "都不扣"
        )));
        RULE_SETS.put("B", new LinkedHashSet<>(Arrays.asList(
                "除招商都扣", "招商", "招商不收", "除招商外都扣", "员工都是招商卡"
        )));
        RULE_SETS.put("C", new LinkedHashSet<>(Arrays.asList(
                "除招商、建设外都扣，离职员工不扣，前15个员工不扣", "除招商和建设外都扣",
                "招商、建设不收", "除招商、建设外都扣",
                "招商、建设不扣，其他银行都扣"
        )));
        RULE_SETS.put("D", new LinkedHashSet<>(Collections.singletonList("除平安外都扣")));
        RULE_SETS.put("E", new LinkedHashSet<>(Collections.singletonList("除宁波外都扣")));
        RULE_SETS.put("F", new LinkedHashSet<>(Collections.singletonList("除工商都扣")));
        RULE_SETS.put("G", new LinkedHashSet<>(Arrays.asList(
                "除温州银行外都扣手续费", "除温州银行外都收手续费"
        )));
        RULE_SETS.put("H", new LinkedHashSet<>(Arrays.asList(
                "邮政卡不收手续费（除邮政都收）", "除了邮政之外的手续费员工承担"
        )));
        RULE_SETS.put("I", new LinkedHashSet<>(Collections.singletonList("除邮政银行，其他银行都扣")));
        RULE_SETS.put("J", new LinkedHashSet<>(Arrays.asList(
                "招商和邮政不扣手续费", "除招商、邮政外都扣", "招商、邮储不扣，其他银行都扣"
        )));
        RULE_SETS.put("K", new LinkedHashSet<>(Collections.singletonList("除邮政，中信外都扣")));
        RULE_SETS.put("L", new LinkedHashSet<>(Collections.singletonList("中信不扣，其他都扣")));
        RULE_SETS.put("M", new LinkedHashSet<>(Collections.singletonList(
                "签浙江锦途：招商、建设不扣；签温州驰锦：民生不扣"
        )));
    }

    // ============================================================
    // 数据模型
    // ============================================================

    /** 汇总结果（对应 Python run_merge_collect 返回的 dict） */
    public static class MergeResult {
        public List<Object> header;
        public List<List<Object>> rows;
        public Integer targetIdCol;          // 身份证列在 header 中的 1-based 位置
        public int totalFiles;
        public int totalRows;
        public int skipped;
        public List<String> errors;
        public Set<String> emptyKeywordOps;
        public Set<String> missingOps;

        public MergeResult() {
            this.header = null;
            this.rows = new ArrayList<>();
            this.targetIdCol = null;
            this.totalFiles = 0;
            this.totalRows = 0;
            this.skipped = 0;
            this.errors = new ArrayList<>();
            this.emptyKeywordOps = new LinkedHashSet<>();
            this.missingOps = new LinkedHashSet<>();
        }
    }

    /** 单文件读取结果（对应 Python _read_payroll_file 返回的 tuple） */
    private static class PayrollData {
        List<Object> srcHeaders;
        List<List<Object>> srcRows;
        int colName;      // 1-based
        int colId;        // 1-based
        int colGjj;       // 1-based
        int colPay;       // 1-based
        Map<String, List<Object[]>> infoLookup;
        boolean infoFound;
        List<String> warns;
    }

    // ============================================================
    // 文本工具函数（对应 Python remove_spaces / normalize_name / clean_project_name / extract_month_from_filename）
    // ============================================================

    /**
     * 去除半角空格与全角空格（与 Python remove_spaces 等价，仅去空格不含其他字符）
     */
    public static String removeSpaces(String s) {
        if (s == null) return "";
        // Python: re.sub(r"[\u3000 ]", "", str(s))
        String result = s.replace("\u3000", "");
        result = result.replace(" ", "");
        return result;
    }

    /**
     * 姓名归一化：去空格，全角数字/字母转半角（与 Python normalize_name 等价）
     */
    public static String normalizeName(String s) {
        if (s == null) return "";
        s = s.trim();
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            // Python: if ch == " " or ch == "\u3000": continue
            if (ch == ' ' || ch == '\u3000') continue;
            int o = (int) ch;
            // Python: if 0xFF10 <= o <= 0xFF19: out.append(chr(o - 0xFF10 + 0x30))
            if (o >= 0xFF10 && o <= 0xFF19) {
                sb.append((char) (o - 0xFF10 + 0x30));
            }
            // Python: elif 0xFF21 <= o <= 0xFF3A: out.append(chr(o - 0xFF21 + 0x41))
            else if (o >= 0xFF21 && o <= 0xFF3A) {
                sb.append((char) (o - 0xFF21 + 0x41));
            }
            // Python: elif 0xFF41 <= o <= 0xFF5A: out.append(chr(o - 0xFF41 + 0x61))
            else if (o >= 0xFF41 && o <= 0xFF5A) {
                sb.append((char) (o - 0xFF41 + 0x61));
            }
            // Python: else: out.append(ch)
            else {
                sb.append(ch);
            }
        }
        return sb.toString();
    }

    /**
     * 清洗"项目"长文本：去年月/工资表/单位元等（与 Python clean_project_name 等价）
     */
    public static String cleanProjectName(String raw) {
        if (raw == null) return "";
        String s = raw.trim();
        if (s.isEmpty()) return "";
        // Python: s = re.sub(r"^\d{2,4}年\d{1,2}月份?", "", s)
        s = s.replaceAll("^\\d{2,4}年\\d{1,2}月份?", "");
        // Python: s = re.sub(r"工资表", "", s)
        s = s.replace("工资表", "");
        // Python: s = re.sub(r"单位：元", "", s)
        s = s.replace("单位：元", "");
        // Python: s = re.sub(r"\s*（", "（", s)
        s = s.replaceAll("\\s*（", "（");
        // Python: s = re.sub(r"\s*\\(", "(", s)
        s = s.replaceAll("\\s*\\(", "(");
        // Python: s = re.sub(r"\s+", " ", s)
        s = s.replaceAll("\\s+", " ");
        return s.trim();
    }

    /**
     * 从文件名提取月份数字（与 Python extract_month_from_filename 等价）
     */
    public static String extractMonthFromFilename(String fname) {
        // Python: m = re.search(r"(\d+)月", fname)
        Matcher m = Pattern.compile("(\\d+)月").matcher(fname);
        if (!m.find()) return "";
        return m.group(1);
    }

    // ============================================================
    // 手续费规则分类与计算（对应 Python classify_rule_exact / is_bank_exempt /
    // resolve_amount / handle_dual_unit_rule / compute_fee_v2）
    // ============================================================

    /**
     * 将规则文本精确匹配到类别 A~M；未命中返回空串（与 Python classify_rule_exact 等价）
     */
    public static String classifyRuleExact(String ruleText) {
        // Python: s = remove_spaces(rule_text)
        String s = removeSpaces(ruleText);
        for (Map.Entry<String, Set<String>> entry : RULE_SETS.entrySet()) {
            if (entry.getValue().contains(s)) {
                return entry.getKey();
            }
        }
        return "";
    }

    /**
     * 按类别 + 开户行（包含匹配）判断是否免扣（与 Python is_bank_exempt 等价）
     */
    public static boolean isBankExempt(String category, String bankName) {
        // Python: b = remove_spaces(bank_name)
        String b = removeSpaces(bankName);
        switch (category) {
            case "A":
                return true;
            case "B":
                return b.contains("招商银行");
            case "C":
                return b.contains("招商银行") || b.contains("建设银行");
            case "D":
                return b.contains("平安银行");
            case "E":
                return b.contains("宁波银行");
            case "F":
                return b.contains("工商");
            case "G":
                return b.contains("温州银行");
            case "H":
            case "I":
                return b.contains("邮政") || b.contains("邮储");
            case "J":
                return b.contains("招商银行") || b.contains("邮政") || b.contains("邮储");
            case "K":
                return b.contains("邮政") || b.contains("邮储") || b.contains("中信银行");
            case "L":
                return b.contains("中信银行");
            default:
                return false;
        }
    }

    /**
     * 解析最终金额：表格有效正数优先，否则按类别默认（A=0, I=3, 其余=5）（与 Python resolve_amount 等价）
     */
    public static int resolveAmount(Object tableAmount, String category) {
        // Python: if table_amount is not None and table_amount != "": try: num = float(table_amount) ...
        if (tableAmount != null && !tableAmount.toString().trim().isEmpty()) {
            try {
                double num = Double.parseDouble(tableAmount.toString().trim());
                if (num > 0) return (int) num;
            } catch (NumberFormatException ignored) {
            }
        }
        if ("A".equals(category)) return 0;
        if ("I".equals(category)) return 3;
        return 5;
    }

    /**
     * M 类双单位特殊规则（与 Python handle_dual_unit_rule 等价）
     */
    public static int handleDualUnitRule(String payUnit, String bank, Object tableAmount) {
        // Python: pu = remove_spaces(pay_unit); b = remove_spaces(bank)
        String pu = removeSpaces(payUnit);
        String b = removeSpaces(bank);
        int amt = 5;
        if (tableAmount != null && !tableAmount.toString().trim().isEmpty()) {
            try {
                double tn = Double.parseDouble(tableAmount.toString().trim());
                if (tn > 0) amt = (int) tn;
            } catch (NumberFormatException ignored) {
            }
        }
        // Python: if pu == "浙江锦途人力资源有限公司": return 0 if ("招商银行" in b or "建设银行" in b) else amt
        if ("浙江锦途人力资源有限公司".equals(pu)) {
            return (b.contains("招商银行") || b.contains("建设银行")) ? 0 : amt;
        }
        // Python: if pu == "温州驰锦企业服务有限公司": return 0 if "民生" in b else amt
        if ("温州驰锦企业服务有限公司".equals(pu)) {
            return b.contains("民生") ? 0 : amt;
        }
        // Python: 兜底: 招商/建行不扣
        return (b.contains("招商银行") || b.contains("建设银行")) ? 0 : amt;
    }

    /**
     * 姓名/身份证无效判定：空 / 数值0 / "0" / "0.0"等 / 错误值（与 Python _is_invalid_name_id 等价）
     *
     * 注意：POI 对数值型单元格返回 double，Excel 中值为 0 时 POI 返回 0.0，
     *       其 toString() 为 "0.0" 而非 "0"，需额外处理数值零值。
     */
    private static boolean isInvalidNameId(Object v) {
        if (v == null) return true;
        String s = v.toString().trim();
        if (s.isEmpty()) return true;
        // 检查数值零值（兼容 POI double 0.0 → "0.0"、以及 "0"、"0.00" 等所有形式）
        try {
            if (Double.parseDouble(s) == 0.0) return true;
        } catch (NumberFormatException ignored) {
            // 非数字字符串，继续后续判断
        }
        // Python: if s.startswith("#"): return True
        if (s.startsWith("#")) return true;
        return false;
    }

    /**
     * 逐行计算手续费（与 Python compute_fee_v2 等价）
     *
     * @param nameVal     姓名单元格值
     * @param idVal       身份证单元格值
     * @param payUnitVal  发薪单位单元格值
     * @param infoLookup  dict{ normalized_name: [(ops, bank), ...] }
     * @param feeRules    dict{ 运营项目: [扣费银行规则文本, 扣费金额] }
     * @param missingOps  可选 set，用于收集"源文件有但规则表缺"的运营项目
     * @return 0 / 3 / 5 / 其他金额，或 null（留空）
     */
    public static Integer computeFeeV2(Object nameVal, Object idVal, Object payUnitVal,
                                        Map<String, List<Object[]>> infoLookup,
                                        Map<String, Object[]> feeRules,
                                        Set<String> missingOps) {
        // Python: if _is_invalid_name_id(name_val) or _is_invalid_name_id(id_val): return None
        if (isInvalidNameId(nameVal) || isInvalidNameId(idVal)) return null;
        // Python: if not fee_rules: return None
        if (feeRules == null || feeRules.isEmpty()) return null;
        // Python: if not info_lookup: return None
        if (infoLookup == null || infoLookup.isEmpty()) return null;

        // Python: key = normalize_name(name_val)
        String key = normalizeName(nameVal != null ? nameVal.toString() : "");
        // Python: entries = info_lookup.get(key)
        List<Object[]> entries = infoLookup.get(key);
        // Python: if not entries: return None
        if (entries == null || entries.isEmpty()) return null;

        for (Object[] entry : entries) {
            // Python: ops, bank in entries
            Object ops = entry[0];
            Object bank = entry[1];
            // Python: if ops is None: continue
            if (ops == null) continue;
            // Python: base_ops = str(ops).strip()
            // 注意：POI 把整数单元格读成 double(106.0)，openpyxl 读成 int(106)→"106"，
            // 为与 Python 行为一致，数值型统一去掉 ".0"（见 normalizeOpsKey）
            String baseOps = normalizeOpsKey(ops);
            // Python: if base_ops == "": continue
            if (baseOps.isEmpty()) continue;
            // Python: if base_ops.startswith("代发"): base_ops = base_ops[2:]
            if (baseOps.startsWith("代发")) baseOps = baseOps.substring(2);
            // Python: if base_ops not in fee_rules: missing_ops.add(base_ops); continue
            if (!feeRules.containsKey(baseOps)) {
                if (missingOps != null) missingOps.add(baseOps);
                continue;
            }
            // Python: rule_text, table_amount = fee_rules[base_ops]
            Object[] rule = feeRules.get(baseOps);
            String ruleText = rule[0] != null ? rule[0].toString() : "";
            Object tableAmount = rule.length > 1 ? rule[1] : null;
            // Python: category = classify_rule_exact(rule_text)
            String category = classifyRuleExact(ruleText);
            // Python: if category == "": continue
            if (category.isEmpty()) continue;
            // Python: if category == "M": return handle_dual_unit_rule(...)
            if ("M".equals(category)) {
                String payUnitStr = payUnitVal != null ? payUnitVal.toString() : "";
                String bankStr = bank != null ? bank.toString() : "";
                return handleDualUnitRule(payUnitStr, bankStr, tableAmount);
            }
            // Python: if is_bank_exempt(category, bank): return 0
            if (isBankExempt(category, bank != null ? bank.toString() : "")) return 0;
            // Python: return resolve_amount(table_amount, category)
            return resolveAmount(tableAmount, category);
        }
        // Python: return None
        return null;
    }

    // ============================================================
    // 规则表加载（对应 Python load_fee_rules）
    // ============================================================

    /**
     * 从规则 Excel 加载规则字典（与 Python load_fee_rules 等价）
     *
     * 自动选择名字含"手续费"的工作表，否则取第一个；
     * 动态定位 客户状态/运营项目/扣费银行（多填）/扣费金额 列；
     * 跳过"客户状态=终止合作"；重复运营项目以最后一条为准。
     *
     * @return dict{ 运营项目: [扣费银行规则文本, 扣费金额] }
     */
    public Map<String, Object[]> loadFeeRules(File file, Consumer<String> onLog) {
        if (onLog == null) onLog = m -> {};
        Map<String, Object[]> rules = new LinkedHashMap<>();

        // Python: wb = load_workbook(path, data_only=True, read_only=True)
        try (Workbook wb = WorkbookFactory.create(new FileInputStream(file))) {
            // Python: 自动选择名字含"手续费"的工作表
            String sheetName = null;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String sn = wb.getSheetName(i);
                if (sn.contains("手续费")) {
                    sheetName = sn;
                    break;
                }
            }
            if (sheetName == null) sheetName = wb.getSheetName(0);
            onLog.accept("规则表使用工作表: " + sheetName);

            Sheet ws = wb.getSheet(sheetName);
            // Python: rows = list(ws.iter_rows(values_only=True))
            List<Object[]> rows = readSheetRows(ws);
            if (rows.isEmpty()) return rules;

            // Python: 动态定位列
            Object[] header = rows.get(0);
            int colStatus = 0, colOps = 0, colBank = 0, colAmount = 0;
            for (int i = 0; i < header.length; i++) {
                String ht = removeSpaces(header[i] != null ? header[i].toString() : "");
                if ("客户状态".equals(ht)) colStatus = i + 1;      // 1-based
                else if ("运营项目".equals(ht)) colOps = i + 1;
                else if ("扣费银行（多填）".equals(ht) || "扣费银行(多填)".equals(ht)) colBank = i + 1;
                else if ("扣费金额".equals(ht)) colAmount = i + 1;
            }

            // Python: if col_ops == 0 or col_bank == 0: return {}
            if (colOps == 0 || colBank == 0) {
                onLog.accept("⚠ 规则表缺少「运营项目」或「扣费银行（多填）」列, 无法加载");
                return rules;
            }

            // Python: for r in rows[1:]: 逐行读取，跳过终止合作
            for (int r = 1; r < rows.size(); r++) {
                Object[] row = rows.get(r);
                // 跳过"客户状态=终止合作"
                if (colStatus > 0) {
                    Object st = colStatus - 1 < row.length ? row[colStatus - 1] : null;
                    if (st != null && "终止合作".equals(removeSpaces(st.toString()))) continue;
                }
                // 运营项目为空则跳过
                Object opsVal = colOps - 1 < row.length ? row[colOps - 1] : null;
                if (opsVal == null || opsVal.toString().trim().isEmpty()) continue;
                String opsKey = normalizeOpsKey(opsVal);
                // 扣费银行规则文本
                Object bankVal = colBank - 1 < row.length ? row[colBank - 1] : null;
                // 扣费金额
                Object amountVal = (colAmount > 0 && colAmount - 1 < row.length) ? row[colAmount - 1] : null;
                // Python: rules[ops_key] = [bank_val, amount_val]
                rules.put(opsKey, new Object[]{bankVal, amountVal});
            }

            onLog.accept("已加载手续费规则 " + rules.size() + " 条");
        } catch (Exception e) {
            onLog.accept("⚠ 读取规则表失败: " + e.getMessage());
            log.error("加载规则表失败", e);
        }
        return rules;
    }

    // ============================================================
    // 扫描 Excel 文件（对应 Python scan_excel_files）
    // ============================================================

    /**
     * 扫描 Excel 文件列表（与 Python scan_excel_files 等价）
     *
     * @param root         扫描根目录
     * @param recursive    是否递归子目录（始终 true，与 Python run_merge_collect 一致）
     * @param excludePaths 排除的路径集合
     */
    public static List<File> scanExcelFiles(File root, boolean recursive, Set<String> excludePaths) {
        List<File> result = new ArrayList<>();
        Set<String> exclude = new LinkedHashSet<>();
        if (excludePaths != null) {
            for (String p : excludePaths) {
                try {
                    exclude.add(new File(p).getCanonicalPath());
                } catch (IOException e) {
                    exclude.add(new File(p).getAbsolutePath());
                }
            }
        }

        if (recursive) {
            scanRecursive(root, exclude, result);
        } else {
            File[] files = root.listFiles();
            if (files != null) {
                for (File f : files) {
                    if (f.isFile() && !shouldSkip(f, exclude)) {
                        result.add(f);
                    }
                }
            }
        }
        return result;
    }

    private static void scanRecursive(File dir, Set<String> exclude, List<File> result) {
        File[] files = dir.listFiles();
        if (files == null) return;
        for (File f : files) {
            if (f.isDirectory()) {
                scanRecursive(f, exclude, result);
            } else if (f.isFile() && !shouldSkip(f, exclude)) {
                result.add(f);
            }
        }
    }

    private static boolean shouldSkip(File f, Set<String> exclude) {
        // Python: if fn.startswith("~$"): return True
        String fn = f.getName();
        if (fn.startsWith("~$")) return true;
        // Python: if not fn.lower().endswith((".xlsx", ".xls")): return True
        String lower = fn.toLowerCase();
        if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) return true;
        // Python: if os.path.normcase(os.path.abspath(full)) in exclude: return True
        try {
            if (exclude.contains(f.getCanonicalPath())) return true;
        } catch (IOException e) {
            if (exclude.contains(f.getAbsolutePath())) return true;
        }
        return false;
    }

    // ============================================================
    // 读取单个个税工作簿（对应 Python _read_payroll_file）
    // ============================================================

    /**
     * 读取单个工作簿的个税统计表 + 员工信息表（与 Python _read_payroll_file 等价）
     *
     * @return PayrollData，若无"个税统计"表返回 null
     */
    private PayrollData readPayrollFile(File file) throws Exception {
        // Python: wb = load_workbook(fpath, data_only=True, read_only=True)
        try (Workbook wb = WorkbookFactory.create(new FileInputStream(file))) {
            // ---- 找"个税统计"表 ----
            // Python: for sn in wb.sheetnames: if remove_spaces(sn) == "个税统计": target = sn
            Sheet targetSheet = null;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                if ("个税统计".equals(removeSpaces(wb.getSheetName(i)))) {
                    targetSheet = wb.getSheetAt(i);
                    break;
                }
            }
            // Python: if target is None: wb.close(); return None
            if (targetSheet == null) return null;

            // Python: rows = list(ws.iter_rows(values_only=True))
            List<Object[]> rows = readSheetRows(targetSheet);
            if (rows.isEmpty()) return null;

            // ---- 找"员工信息-X月"表 ----
            // 直接查询本工作簿内的"员工信息-"表，以文件内表名为准，不依赖文件名月份
            // （修复：文件名月份可能与表内实际月份不一致，导致命中空/错位表使 infoLookup 为空）
            String infoSheetName = null;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String sn = wb.getSheetName(i);
                if (removeSpaces(sn).startsWith("员工信息-")) {
                    infoSheetName = sn;
                    break;
                }
            }

            // ---- 读取员工信息表构建 info_lookup ----
            Map<String, List<Object[]>> infoLookup = new LinkedHashMap<>();
            boolean infoFound = infoSheetName != null;
            List<String> warns = new ArrayList<>();

            if (infoFound) {
                Sheet infoWs = wb.getSheet(infoSheetName);
                List<Object[]> infoRows = readSheetRows(infoWs);
                if (!infoRows.isEmpty()) {
                    Object[] iheader = infoRows.get(0);
                    int ci = 0, co = 0, cb = 0;
                    for (int i = 0; i < iheader.length; i++) {
                        String ht = removeSpaces(iheader[i] != null ? iheader[i].toString() : "");
                        if ("姓名".equals(ht)) ci = i + 1;
                        else if ("运营项目".equals(ht)) co = i + 1;
                        else if ("开户行".equals(ht)) cb = i + 1;
                    }
                    if (ci > 0 && co > 0 && cb > 0) {
                        for (int r = 1; r < infoRows.size(); r++) {
                            Object[] row = infoRows.get(r);
                            Object nm = ci - 1 < row.length ? row[ci - 1] : null;
                            if (nm == null || nm.toString().trim().isEmpty()) continue;
                            Object ops = co - 1 < row.length ? row[co - 1] : null;
                            Object bank = cb - 1 < row.length ? row[cb - 1] : null;
                            // Python: info_lookup.setdefault(normalize_name(nm), []).append((ops, bank))
                            infoLookup.computeIfAbsent(normalizeName(nm.toString()), k -> new ArrayList<>())
                                    .add(new Object[]{ops, bank});
                        }
                    } else {
                        warns.add("⚠ 员工信息表缺少 姓名/运营项目/开户行 列");
                    }
                }
            }

            // ---- 主表处理 ----
            // Python: header = list(rows[0])
            Object[] header = rows.get(0);

            // Python: 实际列数（最后有数据的列）
            int srcCols = 0;
            for (Object[] row : rows) {
                for (int j = row.length - 1; j >= 0; j--) {
                    if (row[j] != null && !row[j].toString().trim().isEmpty()) {
                        if (j + 1 > srcCols) srcCols = j + 1;
                        break;
                    }
                }
            }
            if (srcCols == 0) srcCols = header.length;

            // ---- 定位关键列 ----
            // Python: col_name = col_id = col_gjj = col_pay = 0
            int colName = 0, colId = 0, colGjj = 0, colPay = 0;
            for (int i = 0; i < Math.min(header.length, srcCols); i++) {
                String ht = removeSpaces(header[i] != null ? header[i].toString() : "");
                if ("姓名".equals(ht)) colName = i + 1;
                else if ("身份证".equals(ht)) colId = i + 1;
                else if ("公积金".equals(ht)) colGjj = i + 1;
                else if ("发薪单位".equals(ht)) colPay = i + 1;
            }

            // Python: if col_gjj == 0: raise ValueError("表头缺少「公积金」列")
            if (colGjj == 0) {
                throw new IllegalArgumentException("表头缺少「公积金」列");
            }
            // Python: if col_name == 0 or col_id == 0: raise ValueError(...)
            if (colName == 0 || colId == 0) {
                throw new IllegalArgumentException("表头缺少「姓名」或「身份证」列");
            }

            // Python: src_headers = [(header[i-1] if i-1 < len(header) else None) for i in range(1, src_cols+1)]
            List<Object> srcHeaders = new ArrayList<>();
            for (int i = 0; i < srcCols; i++) {
                srcHeaders.add(i < header.length ? header[i] : null);
            }

            // Python: src_rows = [] for r in rows[1:]: ...
            List<List<Object>> srcRows = new ArrayList<>();
            for (int r = 1; r < rows.size(); r++) {
                Object[] row = rows.get(r);
                List<Object> rr = new ArrayList<>();
                for (int c = 0; c < srcCols; c++) {
                    rr.add(c < row.length ? row[c] : null);
                }
                srcRows.add(rr);
            }

            PayrollData data = new PayrollData();
            data.srcHeaders = srcHeaders;
            data.srcRows = srcRows;
            data.colName = colName;
            data.colId = colId;
            data.colGjj = colGjj;
            data.colPay = colPay;
            data.infoLookup = infoLookup;
            data.infoFound = infoFound;
            data.warns = warns;
            return data;
        }
    }

    // ============================================================
    // 写出汇总表（对应 Python write_merge_output）
    // ============================================================

    /**
     * 写出汇总表，身份证列文本化（与 Python write_merge_output 等价）
     */
    public void writeMergeOutput(File outputFile, List<Object> header, List<List<Object>> allRows,
                                  Integer targetIdCol) throws IOException {
        // Python: wb = Workbook()
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet ws = wb.createSheet();

            // 预创建可复用的单元格样式（避免每单元格创建一个导致 XLSX 样式数溢出）
            DataFormat df = wb.createDataFormat();
            CellStyle intStyle = wb.createCellStyle();     // 整数格式：0 → 显示 "0" 而非 "0.0"
            intStyle.setDataFormat(df.getFormat("0"));
            CellStyle dateStyle = wb.createCellStyle();     // 日期格式：yyyy/m/d
            dateStyle.setDataFormat(df.getFormat("yyyy/m/d"));
            CellStyle textStyle = wb.createCellStyle();     // 文本格式：@
            textStyle.setDataFormat(df.getFormat("@"));

            // Python: ws.append(header)
            Row headerRow = ws.createRow(0);
            for (int i = 0; i < header.size(); i++) {
                Cell cell = headerRow.createCell(i);
                Object val = header.get(i);
                if (val != null) cell.setCellValue(val.toString());
            }

            // Python: for out in all_rows: row_copy = list(out); ...
            for (int r = 0; r < allRows.size(); r++) {
                Row row = ws.createRow(r + 1);
                List<Object> srcRow = new ArrayList<>(allRows.get(r));
                // Python: 身份证列文本化
                if (targetIdCol != null && targetIdCol - 1 < srcRow.size()) {
                    Object v = srcRow.get(targetIdCol - 1);
                    srcRow.set(targetIdCol - 1, v == null ? "" : formatNumberValue(v));
                }
                for (int c = 0; c < srcRow.size(); c++) {
                    Object val = srcRow.get(c);
                    // null 值跳过不创建单元格（Excel 显示为空白）
                    if (val == null) continue;
                    Cell cell = row.createCell(c);
                    if (val instanceof Number) {
                        double d = ((Number) val).doubleValue();
                        // -1 是 Excel 常用的"无值"标记（尤其日期列），写为空白
                        if (d == -1.0) continue;
                        cell.setCellValue(d);
                        // 整数值套整数格式，显示为 "0"/"5" 而非 "0.0"/"5.0"，与 VBA 一致
                        if (d == (long) d && Math.abs(d) < 1e15) {
                            cell.setCellStyle(intStyle);
                        }
                    } else if (val instanceof java.util.Date) {
                        cell.setCellValue((java.util.Date) val);
                        cell.setCellStyle(dateStyle);
                    } else if (val instanceof String) {
                        cell.setCellValue((String) val);
                    } else {
                        cell.setCellValue(val.toString());
                    }
                }
                // 身份证列强制文本格式（覆盖上面可能设置的数字格式）
                if (targetIdCol != null && targetIdCol > 0 && targetIdCol <= srcRow.size()) {
                    Cell idCell = row.getCell(targetIdCol - 1);
                    if (idCell != null) idCell.setCellStyle(textStyle);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // ============================================================
    // 主流程：合并汇总（对应 Python run_merge_collect）
    // ============================================================

    /**
     * 合并个税 - 扫描 + 读取 + 内存汇总（与 Python run_merge_collect 等价）
     *
     * @param dataFiles 所有待处理的 Excel 文件列表（对应 Python scan_excel_files 的结果）
     * @param rulesFile 手续费规则表文件
     * @param onLog     日志回调（对应 Python on_log callback）
     * @return MergeResult 汇总结果（header, rows, targetIdCol, 统计信息等）
     */
    public MergeResult runMergeCollect(List<File> dataFiles, File rulesFile, Consumer<String> onLog) {
        if (onLog == null) onLog = m -> {};

        MergeResult result = new MergeResult();

        // ---- 加载规则（对应 Python load_fee_rules(rules_path, on_log)） ----
        Map<String, Object[]> feeRules;
        try {
            feeRules = loadFeeRules(rulesFile, onLog);
        } catch (Exception e) {
            onLog.accept("⚠ 读取规则表失败: " + e.getMessage());
            return result;
        }
        if (feeRules.isEmpty()) {
            onLog.accept("⚠ 未加载到任何手续费规则, 手续费列将全部留空");
        }

        // ---- 校验：规则表中"扣费银行（多填）"列是否为空（对应 Python gEmptyKeywordOps） ----
        Set<String> emptyKeywordOps = result.emptyKeywordOps;
        for (Map.Entry<String, Object[]> entry : feeRules.entrySet()) {
            Object ruleText = entry.getValue()[0];
            if (ruleText == null || removeSpaces(ruleText.toString()).isEmpty()) {
                emptyKeywordOps.add(entry.getKey());
            }
        }

        if (dataFiles.isEmpty()) {
            onLog.accept("未找到任何 Excel 文件");
            return result;
        }

        result.totalFiles = dataFiles.size();
        Set<String> missingOps = result.missingOps;

        // ---- 首文件模板 + 按列名映射（对应 Python 同名逻辑） ----
        // 输出表头完全由首个有效文件决定，后续文件按归一化列名做 name-based 映射
        List<Object> templateHeaders = null;
        int templateGjj = 0;
        // 收集每个文件的数据（对应 Python file_data）
        List<FileDataEntry> fileDataList = new ArrayList<>();

        for (int idx = 0; idx < dataFiles.size(); idx++) {
            File f = dataFiles.get(idx);
            String fname = f.getName();
            onLog.accept(String.format("读取中: %s (%d/%d)", fname, idx + 1, dataFiles.size()));

            try {
                PayrollData data = readPayrollFile(f);
                if (data == null) {
                    String msg = "✗ " + fname + " 无「个税统计」表, 已跳过";
                    onLog.accept(msg);
                    result.errors.add(msg);
                    result.skipped++;
                    continue;
                }
                for (String w : data.warns) {
                    onLog.accept(w);
                }
                if (!data.infoFound) {
                    onLog.accept("⚠ " + fname + " 无员工信息表, 手续费将留空");
                }

                // 首个有效文件 → 锁定模板
                if (templateHeaders == null) {
                    templateHeaders = data.srcHeaders;
                    templateGjj = data.colGjj;
                }

                fileDataList.add(new FileDataEntry(f, data));

            } catch (Exception e) {
                String msg = "✗ 无法读取 " + fname + ": " + e.getMessage();
                onLog.accept(msg);
                result.errors.add(msg);
                result.skipped++;
            }
        }

        if (fileDataList.isEmpty() || templateHeaders == null) {
            onLog.accept("没有有效数据可汇总！");
            return result;
        }

        // ---- 构建输出表头（与 Python 一致） ----
        // 文件名/项目 + 首文件除第1列外的列（公积金后插手续费） + 采用/不采用
        // 关键列 → 稳定 key：有列名用归一名，无列名用位置占位 __pos_N__
        boolean feeHeaderInserted = false;
        List<Object> outHeader = new ArrayList<>();
        List<String> outKeys = new ArrayList<>();
        outHeader.add("文件名");
        outKeys.add("_fname");
        outHeader.add("项目");
        outKeys.add("_proj");

        for (int i = 1; i < templateHeaders.size(); i++) {
            outHeader.add(templateHeaders.get(i));
            outKeys.add(colKey(templateHeaders.get(i), i));
            if ("公积金".equals(removeSpaces(templateHeaders.get(i) != null ? templateHeaders.get(i).toString() : ""))
                    && !feeHeaderInserted) {
                outHeader.add("手续费");
                outKeys.add("_fee");
                feeHeaderInserted = true;
            }
        }
        outHeader.add("采用/不采用");
        outKeys.add("_tag");

        // 用归一名查找"姓名"/"身份证"在输出表头中的索引
        Integer nameOutIdx = null;
        Integer idOutIdx = null;
        for (int i = 0; i < outHeader.size(); i++) {
            String norm = removeSpaces(outHeader.get(i) != null ? outHeader.get(i).toString() : "");
            if ("姓名".equals(norm)) nameOutIdx = i;
            if ("身份证".equals(norm)) idOutIdx = i;
        }
        result.targetIdCol = idOutIdx != null ? idOutIdx + 1 : null;  // 1-based

        // ---- 逐文件映射数据（对应 Python for fd in file_data:） ----
        List<List<Object>> allRows = new ArrayList<>();

        for (FileDataEntry fd : fileDataList) {
            String shortName = fd.file.getName();
            // Python: fname = os.path.splitext(os.path.basename(fpath))[0]
            int dotIdx = shortName.lastIndexOf('.');
            String fname = dotIdx > 0 ? shortName.substring(0, dotIdx) : shortName;
            PayrollData pd = fd.data;
            List<Object> srcHeaders = pd.srcHeaders;
            List<List<Object>> srcRows = pd.srcRows;
            int colName = pd.colName;
            int colId = pd.colId;
            int colPay = pd.colPay;
            Map<String, List<Object[]>> infoLookup = pd.infoLookup;

            // 本文件：稳定 key → 源行 0-based 索引
            Map<String, Integer> srcKeyToIdx = new LinkedHashMap<>();
            for (int i = 0; i < srcHeaders.size(); i++) {
                srcKeyToIdx.put(colKey(srcHeaders.get(i), i), i);
            }

            for (List<Object> row : srcRows) {
                // Python: proj = clean_project_name(row[0] if len(row) > 0 else None)
                Object proj = row.size() > 0 ? cleanProjectName(row.get(0) != null ? row.get(0).toString() : null) : "";
                List<Object> out = new ArrayList<>();
                out.add(fname);
                out.add(proj);
                boolean feeComputed = false;

                // 遍历除 文件名/项目(前2) 和 采用/不采用(末列) 外的输出列
                for (int i = 2; i < outKeys.size() - 1; i++) {
                    String key = outKeys.get(i);
                    if ("_fee".equals(key)) {
                        if (!feeComputed) {
                            // 计算手续费
                            Object nameV = (colName > 0 && colName - 1 < row.size()) ? row.get(colName - 1) : null;
                            Object idV = (colId > 0 && colId - 1 < row.size()) ? row.get(colId - 1) : null;
                            Object payV = (colPay > 0 && colPay - 1 < row.size()) ? row.get(colPay - 1) : null;
                            Integer fee = computeFeeV2(nameV, idV, payV, infoLookup, feeRules, missingOps);
                            out.add(fee);
                            feeComputed = true;
                        }
                        continue;
                    }
                    Integer srcIdx = srcKeyToIdx.get(key);
                    Object val = (srcIdx != null && srcIdx < row.size()) ? row.get(srcIdx) : null;
                    out.add(val);
                }

                // 采用/不采用
                Object nv = (nameOutIdx != null && nameOutIdx < out.size()) ? out.get(nameOutIdx) : null;
                Object iv = (idOutIdx != null && idOutIdx < out.size()) ? out.get(idOutIdx) : null;
                String tag = (isInvalidNameId(nv) || isInvalidNameId(iv)) ? "不采用" : "采用";
                out.add(tag);

                allRows.add(out);
            }
        }

        result.header = outHeader;
        result.rows = allRows;
        result.totalRows = allRows.size();
        onLog.accept("✔ 内存汇总完成: 共 " + allRows.size() + " 条数据");

        // ---- 提示汇总①：规则表"扣费银行（多填）"列为空（对应 Python gEmptyKeywordOps） ----
        if (!emptyKeywordOps.isEmpty()) {
            StringBuilder sb = new StringBuilder();
            for (String o : sortSet(emptyKeywordOps)) {
                sb.append("[").append(o).append("] ");
            }
            onLog.accept("【规则表提示】以下运营项目的\"扣费银行（多填）\"列为空，需要填写关键词：" + sb.toString().trim());
        }

        // ---- 提示汇总②：源文件有但规则表缺的运营项目（对应 Python gMissingOps） ----
        if (!missingOps.isEmpty()) {
            StringBuilder sb = new StringBuilder();
            for (String o : sortSet(missingOps)) {
                sb.append("[").append(o).append("] ");
            }
            onLog.accept("【项目提示】以下运营项目在规则表中找不到，请在\"收取银行手续费信息\"工作表中填写运营项目及文本关键词：" + sb.toString().trim());
        }

        return result;
    }

    // ============================================================
    // 辅助方法
    // ============================================================

    /**
     * 运营项目名归一化（用于规则表 key 与员工信息 base_ops 的匹配）
     *
     * 关键差异补偿：openpyxl 把整数值单元格读成 int（106 → "106"），
     * 而 POI 一律读成 double（106.0 → "106.0"）。当规则表与员工信息表的
     * "运营项目"列一个存为文本、一个存为数字时，POI 的 ".0" 后缀会导致
     * containsKey 匹配失败、手续费静默留空。此处对整数值去掉 ".0"，
     * 使 Java 与 Python 的字符串表现一致。对应 Python str(int)/str(float)。
     */
    private static String normalizeOpsKey(Object v) {
        if (v == null) return "";
        // trim 去半角空白 + removeSpaces 去全角空格 \u3000（Excel 从外部系统导入的文本常携带）
        String s = removeSpaces(v.toString().trim());
        if (s.isEmpty()) return "";
        try {
            double d = Double.parseDouble(s);
            if (d == (long) d && Math.abs(d) < 1e15) {
                return Long.toString((long) d);
            }
        } catch (NumberFormatException ignored) {
            // 非数字（如中文项目名），原样返回
        }
        return s;
    }

    /**
     * 数值格式化为与 VBA/openpyxl 一致的表现形式
     * POI 将整数也读为 double（0.0/5.0），此处还原为整数串（"0"/"5"），
     * 浮点保持原样（"0.5"）。对应 Python str(int) 与 str(float) 的差异。
     */
    private static String formatNumberValue(Object v) {
        if (v instanceof Number) {
            double d = ((Number) v).doubleValue();
            if (d == (long) d) {
                return Long.toString((long) d);
            }
            return String.valueOf(d);
        }
        return v.toString();
    }

    /**
     * 列名 → 稳定 key（与 Python _col_key 等价）
     * 有列名用归一名，无列名（空白表头）用位置占位 __pos_N__
     */
    private static String colKey(Object headerText, int pos) {
        String norm = removeSpaces(headerText != null ? headerText.toString() : "");
        return norm.isEmpty() ? "__pos_" + pos + "__" : norm;
    }

    /** 读取工作表所有行数据为 Object[][] */
    private static List<Object[]> readSheetRows(Sheet sheet) {
        List<Object[]> rows = new ArrayList<>();
        int lastRow = sheet.getLastRowNum();
        if (lastRow < 0) return rows;

        // 确定实际最大列数
        int maxCol = 0;
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null) {
                maxCol = Math.max(maxCol, row.getLastCellNum());
            }
        }

        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            Object[] rowData = new Object[maxCol];
            if (row != null) {
                for (int c = 0; c < maxCol; c++) {
                    rowData[c] = readCellValue(row.getCell(c));
                }
            }
            rows.add(rowData);
        }
        return rows;
    }

    /**
     * 读取单元格值，处理公式单元格（对应 openpyxl data_only=True 的行为）
     * 返回：String / Double / Boolean / java.util.Date / null
     *
     * 注意：日期单元格在 Excel 中以数值形式存储，需通过 DateUtil.isCellDateFormatted 判断
     */
    private static Object readCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // 对应 openpyxl data_only=True: 日期自动转为 datetime
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // 对应 openpyxl data_only=True：读取公式缓存结果
                try {
                    CellType cachedType = cell.getCachedFormulaResultType();
                    if (cachedType == CellType.STRING) {
                        return cell.getStringCellValue();
                    } else if (cachedType == CellType.NUMERIC) {
                        // 公式结果也可能是日期
                        if (DateUtil.isCellDateFormatted(cell)) {
                            return cell.getDateCellValue();
                        }
                        return cell.getNumericCellValue();
                    } else if (cachedType == CellType.BOOLEAN) {
                        return cell.getBooleanCellValue();
                    }
                } catch (Exception ignored) {
                    // 缓存不可用时，兜底返回公式字符串，避免静默丢失（如规则文本列）
                    try {
                        return cell.getCellFormula();
                    } catch (Exception ignored2) {
                        return null;
                    }
                }
                return null;
            case ERROR:
            case BLANK:
            default:
                return null;
        }
    }

    /** 对 Set 排序（对应 Python sorted()） */
    private static List<String> sortSet(Set<String> set) {
        List<String> list = new ArrayList<>(set);
        Collections.sort(list);
        return list;
    }

    /** 内部数据类：文件与 PayrollData 的绑定 */
    private static class FileDataEntry {
        final File file;
        final PayrollData data;

        FileDataEntry(File file, PayrollData data) {
            this.file = file;
            this.data = data;
        }
    }
}
