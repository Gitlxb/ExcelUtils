package com.example.project.service.hr.hyp;

import org.apache.poi.xwpf.usermodel.*;
import java.util.*;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 从 {@link WordMergeService} 中抽取的「分类关键词提取」职责。
 *
 * <p>包含三级检测，独立并行、结果合并：
 * <ol>
 *   <li>0-X 分数模式 — 匹配 "0-X分" 行，自动探测 parent 列</li>
 *   <li>主观回退 — 查找"主观"标记行，提取左侧关键词（含顿号拆分）</li>
 *   <li>满分值模式 — 动态探测分数列锚点，向左推断关键词列与 parent 列</li>
 * </ol>
 *
 * <p>所有方法均为无状态静态方法，仅依赖 POI 与
 * {@link WordMergeService#Category} 领域模型，不耦合其余合并逻辑。
 */
public final class CategoryExtractor {

    private CategoryExtractor() {}

    /** 一级检测：0-X 分数模式 */
    public static List<WordMergeService.Category> extractScoreCategories(XWPFDocument doc) {
        Map<String, Set<String>> parentKw = new LinkedHashMap<>();
        Map<XWPFTable, Integer> tableParentCol = new IdentityHashMap<>();
        Pattern lineHasScore = Pattern.compile("[：:]\\s*0-\\d+");
        Pattern scoreLine = Pattern.compile("^\\s*(?:\\d+[、.\\uFF0E]\\s*)?(.+?)[：:]\\s*0-\\d+");

        for (XWPFTable table : doc.getTables()) {
            List<XWPFTableRow> rows = table.getRows();
            for (int ri = 0; ri < rows.size(); ri++) {
                List<XWPFTableCell> cells = rows.get(ri).getTableCells();
                for (int ci = 0; ci < cells.size(); ci++) {
                    String text = getCellText(cells.get(ci));
                    if (text == null || text.trim().isEmpty()) continue;
                    if (!lineHasScore.matcher(text).find()) continue;

                    boolean detected = false;
                    for (String line : text.split("\n")) {
                        line = line.trim();
                        if (line.isEmpty()) continue;
                        Matcher m = scoreLine.matcher(line);
                        if (!m.find()) continue;

                        // ---------- 关键词清洗（与 Python 一致）----------
                        // 1) 先剥尾部标点（含全角句号 。）
                        String kw = m.group(1).trim();
                        kw = kw.replaceAll("[；;、，,。.\\s]+$", "");
                        // 2) 再剥尾部括号说明（锚定行尾，避免误删中间括号）
                        kw = kw.replaceAll("[（(][^）)]*[）)]\\s*$", "").trim();
                        if (kw.isEmpty() || kw.length() < 2) continue;

                        Integer pcol = tableParentCol.get(table);
                        if (pcol == null) {
                            pcol = detectParentCol(cells, ci, false);
                            if (pcol != null) tableParentCol.put(table, pcol);
                        }
                        String parentText = null;
                        if (pcol != null && pcol < cells.size()) {
                            parentText = getCellText(cells.get(pcol));
                            if (parentText != null) parentText = parentText.trim();
                            if (parentText.isEmpty()) parentText = null;
                        }
                        if (parentText == null && pcol != null) {
                            parentText = resolveParentText(table, pcol, ri);
                        }
                        if (parentText == null || parentText.isEmpty()) parentText = kw;

                        parentKw.computeIfAbsent(parentText, k -> new LinkedHashSet<>()).add(kw);
                        detected = true;
                        break;
                    }
                    if (detected) break;
                }
            }
        }
        return toCategories(parentKw, "score");
    }

    /** 二级检测：主观回退 */
    public static List<WordMergeService.Category> extractSubjectiveCategories(XWPFDocument doc) {
        Map<String, Set<String>> parentKw = new LinkedHashMap<>();
        Map<XWPFTable, Integer> colCache = new IdentityHashMap<>();
        Pattern kwPattern = Pattern.compile("^\\s*(?:\\d+[.\\uFF0E]\\d+\\s*)?(.+?)[：:]");

        for (XWPFTable table : doc.getTables()) {
            List<XWPFTableRow> rows = table.getRows();
            for (int ri = 0; ri < rows.size(); ri++) {
                List<XWPFTableCell> cells = rows.get(ri).getTableCells();

                int subjCol = -1;
                for (int ci = 0; ci < cells.size(); ci++) {
                    String t = getCellText(cells.get(ci));
                    if (t != null && t.contains("主观")) { subjCol = ci; break; }
                }
                if (subjCol == -1) continue;

                for (int k = subjCol - 1; k >= 0; k--) {
                    String cellText = getCellText(cells.get(k));
                    if (cellText == null) continue;
                    cellText = cellText.trim();
                    if (cellText.isEmpty()) continue;

                    List<String> bare = new ArrayList<>();
                    for (String line : cellText.split("\n")) {
                        line = line.trim();
                        if (line.isEmpty()) continue;
                        Matcher m = kwPattern.matcher(line);
                        if (m.find()) {
                            String kw = m.group(1).trim();
                            // ---------- 关键词清洗（与 Python 一致）----------
                            kw = kw.replaceAll("[；;、，,。.\\s]+$", "");
                            kw = kw.replaceAll("[（(][^）)]*[）)]\\s*$", "").trim();
                            if (kw.length() >= 2) bare.add(kw);
                        }
                    }
                    if (bare.isEmpty() && cellText.length() > 10 && cellText.contains("、")) {
                        bare.add(cellText);
                    }
                    if (bare.isEmpty()) continue;

                    Integer parentCol = colCache.get(table);
                    if (parentCol == null) {
                        parentCol = detectParentCol(cells, k, true);
                        if (parentCol != null) colCache.put(table, parentCol);
                    }
                    String parentText = null;
                    if (parentCol != null && parentCol < cells.size()) {
                        parentText = getCellText(cells.get(parentCol));
                        if (parentText != null) parentText = parentText.trim();
                        if (parentText.isEmpty()) parentText = null;
                    }
                    if (parentText == null && parentCol != null) {
                        parentText = resolveParentText(table, parentCol, ri);
                    }
                    if (parentText == null || parentText.isEmpty()) parentText = bare.get(0);

                    Set<String> set = parentKw.computeIfAbsent(parentText, x -> new LinkedHashSet<>());
                    for (String bk : bare) {
                        for (String sk : splitByDunhao(bk)) set.add(sk);
                    }
                    break;
                }
            }
        }
        return toCategories(parentKw, "subjective");
    }

    /**
     * 获取单元格文本（段落间用 \n 分隔）。
     * XWPFTableCell.getText() 不插入段落分隔符，会导致多段落内容连成一行，keyword 正则无法按行匹配。
     * 本方法行为与 python-docx 的 cell.text 一致。
     */
    public static String getCellText(XWPFTableCell cell) {
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph p : cell.getParagraphs()) {
            if (sb.length() > 0) sb.append('\n');
            sb.append(p.getText());
        }
        return sb.toString();
    }

    /**
     * 合并单元格向上回退：当前行 parentCol 单元格为空时，向上扫描直到找到非空值。
     */
    public static String resolveParentText(XWPFTable table, int parentCol, int currentRow) {
        if (parentCol < 0) return null;
        for (int pr = currentRow - 1; pr >= 0; pr--) {
            XWPFTableRow prevRow = table.getRow(pr);
            if (prevRow == null) continue;
            List<XWPFTableCell> prevCells = prevRow.getTableCells();
            if (parentCol < prevCells.size()) {
                String t = getCellText(prevCells.get(parentCol));
                if (t != null) {
                    t = t.trim();
                    if (!t.isEmpty()) return t;
                }
            }
        }
        return null;
    }

    public static List<WordMergeService.Category> toCategories(Map<String, Set<String>> parentKw, String type) {
        List<WordMergeService.Category> out = new ArrayList<>();
        for (Map.Entry<String, Set<String>> e : parentKw.entrySet()) {
            WordMergeService.Category c = new WordMergeService.Category();
            c.parent = e.getKey();
            c.type = type;
            c.keywords = new ArrayList<>(e.getValue());
            out.add(c);
        }
        return out;
    }

    /**
     * 自动检测 parent 列（从关键词列向左）。
     * 与 Python _detect_parent_col 一致：
     *   - 排除含冒号 / "0-X" / "主观" / 纯数字分数 / X.Y 章节号 / 纯序号(1、1.) 的单元格
     *   - 取「最靠左、且文字最短」的候选（2~20 字）
     */
    public static Integer detectParentCol(List<XWPFTableCell> cells, int keywordCol, boolean subjective) {
        List<int[]> cands = new ArrayList<>();   // {colIndex, textLength}
        for (int ci = keywordCol - 1; ci >= 0; ci--) {
            String t = getCellText(cells.get(ci));
            if (t == null) continue;
            t = t.trim();
            if (t.isEmpty()) continue;
            if (t.contains("：") || t.contains(":")) continue;
            if (Pattern.compile("0-\\d+").matcher(t).find()) continue;
            if (t.contains("主观")) continue;
            // 纯数字分数：14分 / 6 / 1.4 / 1.2.3
            if (t.matches("^\\s*\\d+(?:[.\\uFF0E]\\d+)*\\s*(分)?\\s*$")) continue;
            // X.Y 章节号：1.4 / 2.5.3
            if (t.matches("^\\s*\\d+[.\\uFF0E]\\d+(?:[.\\uFF0E]\\d+)*\\s*$")) continue;
            // 纯序号：1、 / 1. / 1． （无实际语义）
            if (t.matches("^\\s*\\d+[、.\\uFF0E]\\s*$")) continue;
            // 圆括号序号：(1) / 1)
            if (t.matches("^\\s*[（(]?\\s*\\d+[）)]?\\s*$")) continue;
            // 第X章 / 第X节
            if (t.matches("^\\s*第[一二三四五六七八九十百零\\d]+[章节]?\\s*$")) continue;
            if (t.length() >= 2 && t.length() <= 20) cands.add(new int[]{ci, t.length()});
        }
        if (cands.isEmpty()) return null;
        // 与 Python 一致：文字最短优先，同长取最靠左（colIndex 最小）
        cands.sort((a, b) -> a[1] != b[1] ? Integer.compare(a[1], b[1]) : Integer.compare(a[0], b[0]));
        return cands.get(0)[0];
    }

    /** 顿号拆分（与 Python _split_by_dunhao 一致） */
    public static List<String> splitByDunhao(String text) {
        if (text == null) return new ArrayList<>();
        text = text.trim();
        if (!text.contains("、")) return new ArrayList<>(Collections.singletonList(text));
        String[] parts = text.split("、");
        List<String> result = new ArrayList<>();
        for (int i = 0; i < parts.length; i++) {
            String p = parts[i].trim();
            if (p.isEmpty()) continue;
            if (i == parts.length - 1 && p.contains("，")) {
                p = p.substring(0, p.indexOf("，")).trim();
            }
            if (!p.isEmpty()) result.add(p);
        }
        return result.isEmpty() ? new ArrayList<>(Collections.singletonList(text)) : result;
    }

    /** 4字分块模糊搜索块（与 Python _fuzzy_chunks_4char 一致） */
    public static List<String> fuzzyChunks(String keyword, int minChunk) {
        int n = keyword.length();
        if (n <= minChunk) return new ArrayList<>(Collections.singletonList(keyword));
        List<String> chunks = new ArrayList<>();
        int i = 0;
        while (i < n) { chunks.add(keyword.substring(i, Math.min(i + minChunk, n))); i += minChunk; }
        List<String> valid = new ArrayList<>();
        for (String c : chunks) if (c.length() >= minChunk) valid.add(c);
        if (chunks.size() > 1 && chunks.get(chunks.size() - 1).length() < minChunk) {
            String merged = chunks.get(chunks.size() - 2) + chunks.get(chunks.size() - 1);
            if (!merged.equals(keyword) && !valid.contains(merged)) valid.add(merged);
        }
        List<String> result = new ArrayList<>();
        result.add(keyword);
        for (String c : valid) if (!c.equals(keyword) && !result.contains(c)) result.add(c);
        return result;
    }

    // ===================== 三级检测：满分值模式 =====================

    /**
     * 三级检测（新增）：从"满分值"表格模式中提取分类关键词。
     *
     * <p>锚点：最后一列"满分值"（数字+分）；关键词：分数列左侧最近的非空列；
     * parent：关键词列左侧最近的"短文本列"（平均长度 &lt; 15、无冒号）。
     *
     * <p>特点：动态列探测（不硬编码 col-1/col-2）、关键词经清洗管线去除评分细则/子项/说明。
     */
    public static List<WordMergeService.Category> extractMaxscoreCategories(XWPFDocument doc, Consumer<String> log) {
        Map<String, Set<String>> parentKw = new LinkedHashMap<>();
        List<String> parentOrder = new ArrayList<>();

        List<XWPFTable> tables = doc.getTables();
        for (int tableIdx = 0; tableIdx < tables.size(); tableIdx++) {
            XWPFTable table = tables.get(tableIdx);

            // ---- 阶段 A: 动态探测列位置 ----
            Integer scoreCol = detectScoreColumn(table);
            if (scoreCol == null) continue;

            int[] cols = detectKeywordAndParentColumns(table, scoreCol);
            int keywordCol = cols[0];
            Integer parentCol = cols[1] >= 0 ? cols[1] : null;

            if (keywordCol < 0) {
                if (log != null) log.accept("  [三级·满分值] 表 " + (tableIdx + 1)
                        + ": 找到分数列但未找到关键词列，跳过");
                continue;
            }

            if (log != null) log.accept("  [三级·满分值] 表 " + (tableIdx + 1)
                    + ": score_col=" + scoreCol + ", keyword_col=" + keywordCol
                    + ", parent_col=" + parentCol);

            // ---- 阶段 B: 逐行提取关键词 ----
            String lastParentText = null;
            List<XWPFTableRow> rows = table.getRows();
            for (int ri = 0; ri < rows.size(); ri++) {
                XWPFTableRow row = rows.get(ri);

                if (isTableHeaderRow(row, scoreCol)) {
                    if (log != null) log.accept("  [三级·满分值] 表 " + (tableIdx + 1)
                            + ": 跳过表头行 (ri=" + ri + ")");
                    continue;
                }

                List<XWPFTableCell> cells = row.getTableCells();
                if (keywordCol >= cells.size()) continue;
                String rawText = getCellText(cells.get(keywordCol)).trim();
                if (rawText.isEmpty()) continue;

                String keyword = cleanKeywordText(rawText);
                if (keyword == null || keyword.isEmpty()) continue;

                // 获取 parent
                String parentText = null;
                if (parentCol != null && parentCol < cells.size()) {
                    parentText = getCellText(cells.get(parentCol)).trim();
                }
                if ((parentText == null || parentText.isEmpty()) && lastParentText != null) {
                    parentText = lastParentText;
                }
                if (parentText == null || parentText.isEmpty()) {
                    parentText = keyword;
                }
                // 更新 lastParentText
                if (parentCol != null && parentCol < cells.size()) {
                    String cur = getCellText(cells.get(parentCol)).trim();
                    if (!cur.isEmpty()) lastParentText = cur;
                }

                parentKw.computeIfAbsent(parentText, k -> {
                    parentOrder.add(k);
                    return new LinkedHashSet<>();
                }).add(keyword);

                if (log != null) log.accept("  ✓ [三级·满分值] " + parentText + " ← " + keyword);
            }
        }

        List<WordMergeService.Category> categories = new ArrayList<>();
        for (String parent : parentOrder) {
            Set<String> kws = parentKw.get(parent);
            if (kws == null) continue;
            WordMergeService.Category c = new WordMergeService.Category();
            c.parent = parent;
            c.type = "maxscore";
            c.keywords = new ArrayList<>(kws);
            categories.add(c);
        }

        if (log != null) {
            if (!categories.isEmpty()) {
                int totalKws = categories.stream().mapToInt(c -> c.keywords.size()).sum();
                log.accept("  [三级·满分值] 共 " + categories.size() + " 个类别, " + totalKws + " 个关键词");
            } else {
                log.accept("  [三级·满分值] 未命中");
            }
        }
        return categories;
    }

    /**
     * 判断某行是否为表头行。
     * 策略：任一单元格含典型表头关键词 / 分数列单元格为非数字+分 / 短文本特征匹配。
     */
    public static boolean isTableHeaderRow(XWPFTableRow row, int scoreCol) {
        Set<String> headerKeywords = new HashSet<>(Arrays.asList(
            "满分", "满分值", "分值", "评审项目", "评·分·内·容", "评 分 内 容",
            "评审内容", "评审标准", "评审要求", "评审因素", "评审因素及分值",
            "评分内容", "评分标准", "评分细则", "评分要求", "评分项目",
            "评分因素", "序号", "号", "评审指标", "评审要点"
        ));
        Set<String> headerShortWords = new HashSet<>(Arrays.asList(
            "评审", "评分", "标准", "因素", "指标", "项目"
        ));

        List<XWPFTableCell> cells = row.getTableCells();
        for (int ci = 0; ci < cells.size(); ci++) {
            String text = getCellText(cells.get(ci)).trim();
            if (text.isEmpty()) continue;

            // 完整表头词命中
            if (headerKeywords.contains(text)) return true;

            // 含典型表头关键词
            for (String kw : headerKeywords) {
                if (text.contains(kw) && text.length() <= kw.length() + 2) return true;
            }

            // 分数列单元格：非数字+分格式，含"满分/分值"等表头词
            // Python: any(c in text for c in '满分分值') → 检查 text 是否含这些字符之一
            if (ci == scoreCol) {
                if (!SCORE_VALUE_PATTERN.matcher(text).matches()) {
                    boolean hasHeaderChar = false;
                    for (char hc : "满分分值".toCharArray()) {
                        if (text.indexOf(hc) >= 0) { hasHeaderChar = true; break; }
                    }
                    if (hasHeaderChar || text.length() <= 5) return true;
                }
            }

            // 短文本特征（2-5字，不含正文符号，含表头短词）
            if (text.length() >= 2 && text.length() <= 5) {
                boolean hasPunct = false;
                for (int i = 0; i < text.length(); i++) {
                    char ch = text.charAt(i);
                    if ("：:；;、，。.\n".indexOf(ch) >= 0) { hasPunct = true; break; }
                }
                if (!hasPunct) {
                    for (String sw : headerShortWords) {
                        if (text.contains(sw)) return true;
                    }
                }
            }
        }
        return false;
    }

    /** 满分值检测正则：匹配 "数字.数字分" 格式 */
    private static final Pattern SCORE_VALUE_PATTERN = Pattern.compile("^\\s*\\d+(\\.\\d+)?\\s*分\\s*$");

    /**
     * 动态探测表格的"满分值"列。
     * 统计每列中匹配"数字分"的非空单元格数，取命中最多列；要求命中率 ≥ 50%。
     */
    public static Integer detectScoreColumn(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        int numCols = 0;
        for (XWPFTableRow r : rows) numCols = Math.max(numCols, r.getTableCells().size());
        if (numCols == 0) return null;

        int[] colHits = new int[numCols];
        int[] colNonEmpty = new int[numCols];

        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (int ci = 0; ci < Math.min(cells.size(), numCols); ci++) {
                String text = getCellText(cells.get(ci)).trim();
                if (!text.isEmpty()) {
                    colNonEmpty[ci]++;
                    if (SCORE_VALUE_PATTERN.matcher(text).matches()) {
                        colHits[ci]++;
                    }
                }
            }
        }

        // 取命中最多的列
        int bestCol = -1;
        int maxHits = 0;
        for (int ci = 0; ci < numCols; ci++) {
            if (colHits[ci] > maxHits) {
                maxHits = colHits[ci];
                bestCol = ci;
            }
        }
        if (maxHits == 0) return null;

        // 可靠性检查：≥ 50% 非空行命中
        if (maxHits < colNonEmpty[bestCol] * 0.5) return null;
        return bestCol;
    }

    /**
     * 从分数列向左动态探测关键词列和 parent 列。
     * 返回 int[]{keywordCol, parentCol}，未找到时对应值为 -1。
     */
    public static int[] detectKeywordAndParentColumns(XWPFTable table, int scoreCol) {
        List<XWPFTableRow> rows = table.getRows();

        // 关键词列：分数列左侧最近的非空列
        int keywordCol = -1;
        for (int ci = scoreCol - 1; ci >= 0; ci--) {
            boolean nonEmpty = false;
            for (XWPFTableRow row : rows) {
                if (ci < row.getTableCells().size()) {
                    String t = getCellText(row.getTableCells().get(ci)).trim();
                    if (!t.isEmpty()) { nonEmpty = true; break; }
                }
            }
            if (nonEmpty) { keywordCol = ci; break; }
        }

        // parent 列：关键词列左侧最近的短文本列
        int parentCol = -1;
        if (keywordCol >= 0) {
            for (int ci = keywordCol - 1; ci >= 0; ci--) {
                if (colIsShortText(rows, ci)) { parentCol = ci; break; }
            }
        }

        return new int[]{keywordCol, parentCol};
    }

    /** 判定某列是否为"短文本列"：平均长度 &lt; 15 且不含冒号 */
    private static boolean colIsShortText(List<XWPFTableRow> rows, int ci) {
        List<String> texts = new ArrayList<>();
        for (XWPFTableRow row : rows) {
            if (ci >= row.getTableCells().size()) continue;
            String t = getCellText(row.getTableCells().get(ci)).trim();
            if (!t.isEmpty()) texts.add(t);
        }
        if (texts.isEmpty()) return false;
        double avgLen = texts.stream().mapToInt(String::length).average().orElse(0);
        if (avgLen >= 15) return false;
        for (String t : texts) {
            if (t.contains("：") || t.contains(":")) return false;
        }
        return true;
    }

    /**
     * 满分值模式关键词文本清洗管道。
     *
     * <p>清洗顺序：换行取首段 → 剥离评分细则后缀 → 剥离最高得分说明 →
     * 剥离说明注脚 → 尾部标点/括号去除 → 质量判定（长度≥4且非纯数字）。
     *
     * @return 清洗后的关键词，不可用时返回 null
     */
    public static String cleanKeywordText(String text) {
        if (text == null || text.trim().isEmpty()) return null;
        text = text.trim();

        // ---- 1) 取"第一段核心内容"：冒号+换行+子项列表 → 截取冒号前 ----
        int colonIdx = -1;
        for (String ch : new String[]{"：", ":"}) {
            int idx = text.indexOf(ch);
            if (idx != -1 && (colonIdx == -1 || idx < colonIdx)) colonIdx = idx;
        }
        if (colonIdx != -1) {
            String afterColon = text.substring(colonIdx + 1);
            // Python: r'\s*[\n\r]+\s*[1-9（(][）).、\uFF0E]' — \uFF0E 是全角句号
            if (Pattern.compile("^\\s*[\\n\\r]+\\s*[1-9（(][）).、\uFF0E]").matcher(afterColon).find()
                    || Pattern.compile("^\\s*[\\n\\r]+").matcher(afterColon).find()) {
                text = text.substring(0, colonIdx).trim();
            }
        }

        // ---- 2) 剥离评分细则后缀：由评委打分…。 / 评委打分…。 ----
        // Python: [^\u3002。\n]* → [^。.\n]*
        text = text.replaceAll("[，,]\\s*由评委[^。.\\n]*[。.]", "");
        text = text.replaceAll("[，,]\\s*评委[^。.\\n]*[。.]", "");

        // ---- 3) 剥离最高得分说明：每提供…最高得X分。 / 最高得X分。 ----
        text = text.replaceAll("[，,]\\s*每提供[^。.\\n]*?最高得\\d+分?[。.]", "");
        text = text.replaceAll("[，,]\\s*最高得\\d+分?[。.]", "");

        // ---- 4) 剥离说明注脚：说明：…。 ----
        text = text.replaceAll("[。；;]\\s*说明[：:][^。.]*[。.]", "");

        // ---- 5) 尾部清理（标点 + 括号） ----
        text = text.replaceAll("[；;、，,。]+$", "");
        text = text.replaceAll("[（(][^）)]*[）)]\\s*$", "").trim();

        // ---- 6) 质量判定 ----
        if (text.isEmpty() || text.length() < 4) return null;
        if (Pattern.compile("^\\s*\\d+(\\.[0-9]+)?\\s*分?\\s*$").matcher(text).matches()) return null;

        return text;
    }
}
