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
 * 合并个税：对应 Python hbgs_core.py / VBA「合并个税_修复版.vba」
 * 功能：合并个税统计表，按手续费规则计算银行手续费
 *
 * 核心流程：
 * 1. 加载手续费规则表（目录内「收取手续费银行信息.xlsx」）
 * 2. 逐个读取「个税统计」表 + 「员工信息-X月」表
 * 3. 按 19 列白名单落位（别名归并 / 忽略名单 / 白名单外按邻列插入）
 * 4. 按规则逐行计算手续费（compute_fee_v2）
 * 5. 列尾压缩：「采用/不采用」跟随最后一个有效列
 * 6. 输出汇总表，身份证列文本化（write_merge_output）
 *
 * 2026-10-26 改造（对齐 VBA「按列名映射重写版」）：
 *   原实现为「首文件决定表头 + 按列名映射，多余列丢弃」，与 VBA 旧版同源；
 *   现改为固定白名单 + 动态插列 + 尾部压缩，解决列数 15/16/17 混合批次的丢列问题。
 */
@Service
public class HbgsService {

    private static final Logger log = LoggerFactory.getLogger(HbgsService.class);

    // ============================================================
    // 输出列布局（2026-10-26 对齐 VBA「按列名映射重写版」）
    //   原实现：首文件决定表头 + 按列名映射（多余列丢弃）—— 与 VBA 旧版同源
    //   现实现：固定 19 列白名单 + 白名单外按邻列插入 + 忽略名单 + 尾部压缩
    // ============================================================

    /** 标准输出列白名单（顺序固定，对齐 VBA stdCols，共 19 列） */
    private static final List<String> STD_COLUMNS = Collections.unmodifiableList(Arrays.asList(
            "文件名", "项目", "姓名", "身份证", "入职时间", "离职时间", "应发工资", "个税",
            "预支扣款", "需扣回", "实发工资", "备注", "发薪单位", "联系电话", "社保", "公积金",
            "手续费", "个税申报所属期", "采用/不采用"));

    /**
     * 忽略名单：源表出现这些列名时整列跳过（不进汇总表、不占列位）
     *   · 开户行   —— 手续费引擎从「员工信息-X月」表取，汇总表不需要
     *   · 发放公司 —— 不输出，但 M 类双单位规则仍用它判定发薪单位（见 readPayrollFile 的 colPay）
     *   · 电话号码 —— 与「联系电话」重复，只保留后者
     */
    private static final Set<String> IGNORE_COLUMNS = Collections.unmodifiableSet(
            new LinkedHashSet<>(Arrays.asList("开户行", "发放公司", "电话号码")));

    /** 尾部压缩时必须保留的列（即使整列为空也不裁） */
    private static final Set<String> KEEP_COLUMNS = Collections.unmodifiableSet(
            new LinkedHashSet<>(Arrays.asList("文件名", "项目", "姓名", "身份证", "手续费")));

    // ============================================================
    // 手续费规则分类：A~N 各类精确匹配串（与 VBA 原值逐字一致，含全角标点）
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
        // N-全扣：所有银行都扣手续费（2026-10-26 补齐，此前缺失导致该规则的项目手续费留空）
        RULE_SETS.put("N", new LinkedHashSet<>(Collections.singletonList(
                "所有银行都扣手续费"
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
        public int ruleCount;                // 规则表解析出的规则条数（0 = 无有效规则，调用方应直接失败）
        public List<String> errors;
        public Set<String> emptyKeywordOps;
        public Set<String> missingOps;
        public Set<String> skippedColumns;   // 被忽略名单整列跳过的源列名（去重）

        public MergeResult() {
            this.header = null;
            this.rows = new ArrayList<>();
            this.targetIdCol = null;
            this.totalFiles = 0;
            this.totalRows = 0;
            this.skipped = 0;
            this.ruleCount = 0;
            this.errors = new ArrayList<>();
            this.emptyKeywordOps = new LinkedHashSet<>();
            this.missingOps = new LinkedHashSet<>();
            this.skippedColumns = new LinkedHashSet<>();
        }
    }

    /** 单文件读取结果（对应 Python _read_payroll_file 返回的 tuple） */
    private static class PayrollData {
        List<Object> srcHeaders;
        List<List<Object>> srcRows;
        int srcCols;      // 源表实际列数（最后含值的列）
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
            case "N":
                // 全扣：所有银行都不免除（对齐 VBA IsBankExempt 的 Case "N"）
                return false;
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
                if (ht.isEmpty()) continue;
                if ("姓名".equals(ht)) colName = i + 1;
                else if ("身份证".equals(ht)) colId = i + 1;
                else if ("公积金".equals(ht)) colGjj = i + 1;
                // 发薪单位：先别名归并（发薪公司→发薪单位），再兼容「发放公司」写法。
                // 「发放公司」不进汇总表（忽略名单），但 M 类双单位规则仍需要它判定发薪单位。
                else if (colPay == 0 && ("发薪单位".equals(aliasName(ht)) || "发放公司".equals(ht))) colPay = i + 1;
            }

            // 2026-10-26：公积金列不再强制要求（对齐 VBA —— 缺列留空，不整簿跳过）
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
            int lastNonBlank = -1;
            for (int r = 1; r < rows.size(); r++) {
                Object[] row = rows.get(r);
                List<Object> rr = new ArrayList<>();
                boolean blank = true;
                for (int c = 0; c < srcCols; c++) {
                    Object v = c < row.length ? row[c] : null;
                    rr.add(v);
                    if (blank && v != null && !removeSpaces(v.toString().trim()).isEmpty()) {
                        blank = false;
                    }
                }
                srcRows.add(rr);
                if (!blank) lastNonBlank = srcRows.size() - 1;
            }
            // 裁掉尾部「仅格式无值」的行（对应 VBA srcRows = SheetRow - 1），
            // 否则这些行会以空白数据进入汇总并被标成「不采用」
            if (lastNonBlank + 1 < srcRows.size()) {
                srcRows = new ArrayList<>(srcRows.subList(0, lastNonBlank + 1));
            }

            PayrollData data = new PayrollData();
            data.srcHeaders = srcHeaders;
            data.srcRows = srcRows;
            data.srcCols = srcCols;
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
        result.ruleCount = feeRules.size();
        if (feeRules.isEmpty()) {
            // 2026-10-26：规则表改为目录内唯一来源后，0 条规则即视为失败，不再产出"手续费整列留空"的表
            onLog.accept("⚠ 规则表未解析出任何规则：请检查「运营项目」「扣费银行（多填）」两列是否存在，以及「客户状态」是否全部为「终止合作」");
            return result;
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

        // ============================================================
        // 输出列布局：固定 19 列白名单（对齐 VBA「按列名映射重写版」）
        //   · 源簿缺列 → 该列留空，不影响其他文件
        //   · 白名单外的列 → 按「左邻/右邻」插入输出表并写表头
        //   · 忽略名单（开户行/发放公司/电话号码）→ 整列跳过
        //   · 末尾做「尾部压缩」，「采用/不采用」跟随最后一个有效列
        // ============================================================
        List<String> uNames = new ArrayList<>(STD_COLUMNS);
        Map<String, Integer> uDict = indexOfNames(uNames);
        List<List<Object>> allRows = new ArrayList<>();
        Set<String> skippedColumns = result.skippedColumns;
        int uCount = uNames.size();

        for (int idx = 0; idx < dataFiles.size(); idx++) {
            File f = dataFiles.get(idx);
            onLog.accept(String.format("读取中: %s (%d/%d)", f.getName(), idx + 1, dataFiles.size()));

            PayrollData pd;
            try {
                pd = readPayrollFile(f);
            } catch (Exception e) {
                String msg = "✗ 无法读取 " + f.getName() + ": " + e.getMessage();
                onLog.accept(msg);
                result.errors.add(msg);
                result.skipped++;
                continue;
            }
            if (pd == null) {
                String msg = "✗ " + f.getName() + " 无「个税统计」表, 已跳过";
                onLog.accept(msg);
                result.errors.add(msg);
                result.skipped++;
                continue;
            }
            for (String w : pd.warns) {
                onLog.accept(w);
            }
            if (!pd.infoFound) {
                onLog.accept("⚠ " + f.getName() + " 无员工信息表, 手续费将留空");
            }

            int srcCols = pd.srcCols;

            // ---- ① 表头 → 本地列名：别名归并 + 生成列加「(源)」+ 同簿重名 #2 ----
            String[] loc = new String[srcCols + 1];
            Set<String> usedNames = new LinkedHashSet<>();
            for (int s = 2; s <= srcCols; s++) {
                String nm = localName(s - 1 < pd.srcHeaders.size() ? pd.srcHeaders.get(s - 1) : null);
                if (!nm.isEmpty()) {
                    if (usedNames.contains(nm)) {
                        int c = 2;
                        while (usedNames.contains(nm + "#" + c)) c++;
                        nm = nm + "#" + c;
                    }
                    usedNames.add(nm);
                }
                loc[s] = nm;
            }

            // ---- ② 忽略名单 / 无表头告警 / 白名单外插列 ----
            for (int s = 2; s <= srcCols; s++) {
                String nm = loc[s];
                if (nm.isEmpty()) {
                    // 表头空但列下有值 → 该列数据会丢，明确告警
                    if (columnHasData(pd.srcRows, s - 1)) {
                        onLog.accept("⚠ " + f.getName() + "：第 " + s + " 列无表头，该列数据未汇总");
                    }
                    continue;
                }
                if (IGNORE_COLUMNS.contains(nm)) {
                    skippedColumns.add(nm);
                    continue;
                }
                if (uDict.containsKey(nm)) continue;

                int pos = resolveInsertPos(loc, s, srcCols, uDict, uNames);
                uNames.add(pos - 1, nm);
                uDict = indexOfNames(uNames);
                // 等价 VBA 的 gS2.Columns(pos).Insert：已汇总的行整体右移一格
                for (List<Object> r : allRows) {
                    r.add(pos - 1, null);
                }
                onLog.accept("⚠ " + f.getName() + "：新增列「" + nm + "」插入第 " + pos + " 列");
            }

            // ---- ③ 源列号 → 输出列号 ----
            uCount = uNames.size();
            int[] gIdx = new int[srcCols + 1];
            for (int s = 2; s <= srcCols; s++) {
                if (!loc[s].isEmpty()) {
                    Integer gi = uDict.get(loc[s]);
                    if (gi != null) gIdx[s] = gi;
                }
            }
            int fIdx = uDict.get("文件名");
            int pIdx = uDict.get("项目");
            int idIdx = uDict.get("身份证");
            int feeIdx = uDict.get("手续费");

            String shortName = f.getName();
            int dotIdx = shortName.lastIndexOf('.');
            String fname = dotIdx > 0 ? shortName.substring(0, dotIdx) : shortName;

            // ---- ④ 逐行落位 ----
            for (List<Object> row : pd.srcRows) {
                List<Object> out = new ArrayList<>(Collections.nCopies(uCount, null));
                out.set(fIdx - 1, fname);
                Object projRaw = row.size() > 0 ? row.get(0) : null;
                out.set(pIdx - 1, cleanProjectName(projRaw != null ? projRaw.toString() : null));
                for (int s = 2; s <= srcCols && s - 1 < row.size(); s++) {
                    if (gIdx[s] > 0) {
                        out.set(gIdx[s] - 1, row.get(s - 1));
                    }
                }
                // 身份证文本化（防 18 位变科学计数法）
                out.set(idIdx - 1, idText(out.get(idIdx - 1)));
                // 手续费：引擎按「源表列号」取 姓名 / 身份证 / 发薪单位
                out.set(feeIdx - 1, computeFeeV2(
                        cell(row, pd.colName), cell(row, pd.colId), cell(row, pd.colPay),
                        pd.infoLookup, feeRules, missingOps));
                allRows.add(out);
            }
        }

        if (allRows.isEmpty()) {
            onLog.accept("没有有效数据可汇总！");
            return result;
        }

        // ---- ⑤ 尾部压缩：「采用/不采用」跟随最后一个有效列 ----
        int colsBefore = uNames.size();
        compactTailColumns(uNames, allRows);
        uDict = indexOfNames(uNames);
        uCount = uNames.size();

        // ---- ⑥ 写「采用/不采用」----
        int tagIdx = uDict.get("采用/不采用");
        int nameOutIdx = uDict.get("姓名");
        int idOutIdx = uDict.get("身份证");
        for (List<Object> r : allRows) {
            Object nv = r.get(nameOutIdx - 1);
            Object iv = r.get(idOutIdx - 1);
            r.set(tagIdx - 1, (isInvalidNameId(nv) || isInvalidNameId(iv)) ? "不采用" : "采用");
        }

        result.header = new ArrayList<>(uNames);
        result.rows = allRows;
        result.totalRows = allRows.size();
        result.targetIdCol = uDict.get("身份证");
        onLog.accept("✔ 内存汇总完成: 共 " + allRows.size() + " 条数据, 输出 " + uCount + " 列");
        if (colsBefore != uCount) {
            onLog.accept("列尾压缩：删除尾部 " + (colsBefore - uCount) + " 个空列，「采用/不采用」位于第 " + tagIdx + " 列");
        }

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

        // ---- 提示汇总③：被忽略名单整列跳过的源列（对应 VBA gSkippedCols） ----
        if (!skippedColumns.isEmpty()) {
            StringBuilder sb = new StringBuilder();
            for (String o : sortSet(skippedColumns)) {
                sb.append("[").append(o).append("] ");
            }
            onLog.accept("【已跳过列】以下源表列按忽略名单整列跳过，未汇总到结果表：" + sb.toString().trim());
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

    // ============================================================
    // 列对齐 / 行落位辅助（对齐 VBA AliasName / LocalName / SyncUnion / CompactTailColumns）
    // ============================================================

    /**
     * 源表头 → 标准列名：别名归并（对齐 VBA AliasName）
     *
     * 注意：「发放公司」「电话号码」【不归并】—— 它们要命中 SyncUnion 的忽略名单整列跳过；
     *       若归并成「发薪单位」「联系电话」就跳过不生效了。「发薪公司」「手机号」是另一种写法，仍归并。
     */
    public static String aliasName(String h) {
        if (h == null) return "";
        switch (h) {
            case "个税扣款":
                return "个税";
            case "入职日期":
                return "入职时间";
            case "离职日期":
                return "离职时间";
            case "发薪公司":
                return "发薪单位";
            case "手机号":
                return "联系电话";
            case "预支":
            case "预支扣回":
                return "预支扣款";
            default:
                // 带月份戳的表头（如「个税申报所属期:2026年8月」）按前缀归一
                if (h.startsWith("个税申报所属期")) return "个税申报所属期";
                return h;
        }
    }

    /** 源表某列的表头 → 本地列名；返回 "" 表示无表头（该列跳过） */
    private static String localName(Object hdrVal) {
        if (hdrVal == null) return "";
        String h = removeSpaces(hdrVal.toString().trim());
        if (h.isEmpty()) return "";
        h = aliasName(h);
        // 与生成列同名的源列 → 加「(源)」后缀另立一列，避免覆盖生成结果
        if ("文件名".equals(h) || "项目".equals(h) || "手续费".equals(h) || "采用/不采用".equals(h)) {
            h = h + "(源)";
        }
        return h;
    }

    /** 列名表 → 索引（列名 → 1-based 列号） */
    private static Map<String, Integer> indexOfNames(List<String> names) {
        Map<String, Integer> m = new LinkedHashMap<>();
        for (int i = 0; i < names.size(); i++) {
            m.put(names.get(i), i + 1);
        }
        return m;
    }

    /** 取行内单元格（1-based 列号），越界返回 null */
    private static Object cell(List<Object> row, int col1Based) {
        if (col1Based <= 0 || col1Based > row.size()) return null;
        return row.get(col1Based - 1);
    }

    /** 该列（0-based）在数据行中是否有非空值 */
    private static boolean columnHasData(List<List<Object>> rows, int col0) {
        for (List<Object> r : rows) {
            if (col0 < r.size()) {
                Object v = r.get(col0);
                if (v != null && !removeSpaces(v.toString().trim()).isEmpty()) return true;
            }
        }
        return false;
    }

    /** 身份证文本化：空保持空；数值转全位数字符串（避免 18 位变科学计数法） */
    private static Object idText(Object v) {
        if (v == null) return null;
        if (v instanceof Number) return formatNumberValue(v);
        String s = v.toString();
        return s.isEmpty() ? null : s;
    }

    /**
     * 新列插入位置（对齐 VBA SyncUnion）：
     *   左右邻相邻 → 插中间；不相邻 → 以左邻为准；只有一侧 → 紧邻该侧
     */
    private static int resolveInsertPos(String[] loc, int s, int srcCols,
                                        Map<String, Integer> uDict, List<String> uNames) {
        String pName = "";
        String nName = "";
        for (int j = s - 1; j >= 2; j--) {
            if (loc[j] != null && !loc[j].isEmpty() && uDict.containsKey(loc[j])) {
                pName = loc[j];
                break;
            }
        }
        for (int j = s + 1; j <= srcCols; j++) {
            if (loc[j] != null && !loc[j].isEmpty() && uDict.containsKey(loc[j])) {
                nName = loc[j];
                break;
            }
        }
        if (!pName.isEmpty() && !nName.isEmpty()) {
            int pIdx = uDict.get(pName);
            int nIdx = uDict.get(nName);
            if (nIdx == pIdx + 1) return pIdx + 1;   // 左右邻相邻 → 插中间
            return nIdx > pIdx ? pIdx + 1 : nIdx;    // 列序异常时以较靠前者为准
        }
        if (!pName.isEmpty()) return uDict.get(pName) + 1;
        if (!nName.isEmpty()) return uDict.get(nName);
        // 两侧都没有可定位的列 → 放在标记列之前（标记列恒在末位）
        return "采用/不采用".equals(uNames.get(uNames.size() - 1)) ? uNames.size() : uNames.size() + 1;
    }

    /**
     * 列尾压缩（对齐 VBA CompactTailColumns）：
     *   删掉「最后一个有效列」右侧的空列，使「采用/不采用」紧跟数据。
     *   只裁尾部 —— 中间空列保留，保证跨月列序稳定。
     */
    private static void compactTailColumns(List<String> uNames, List<List<Object>> rows) {
        int tagIdx = uNames.indexOf("采用/不采用") + 1;      // 1-based
        if (tagIdx <= 1 || tagIdx != uNames.size()) return;  // 防御：标记列不在末位时不动手

        int lastData = 0;
        for (int c = tagIdx - 1; c >= 1; c--) {
            String nm = uNames.get(c - 1);
            if (KEEP_COLUMNS.contains(nm)) {
                lastData = c;
                break;
            }
            if (columnHasData(rows, c - 1)) {
                lastData = c;
                break;
            }
        }
        if (lastData == 0 || lastData >= tagIdx - 1) return;  // 无尾部空列 → 零改动

        for (int c = tagIdx - 1; c >= lastData + 1; c--) {
            uNames.remove(c - 1);
            for (List<Object> r : rows) {
                if (c - 1 < r.size()) r.remove(c - 1);
            }
        }
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
}
