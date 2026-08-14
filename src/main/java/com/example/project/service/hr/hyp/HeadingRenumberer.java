package com.example.project.service.hr.hyp;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.*;

import java.math.BigInteger;
import java.util.*;
import java.util.regex.*;

import static com.example.project.service.hr.hyp.OoxmlDomUtils.*;

/**
 * 标题注入与静态编号回填子系统（从 WordMergeService 提取的 D+F 聚类）。
 *
 * 职责：
 *   - 在合并文档中插入「一级分类标题」（insertCategoryHeading），其样式/编号/outlineLvl 落地正确；
 *   - 定位（或动态注入）基础文档的一级标题样式 ID（findHeading1Id / injectHeading1Style）；
 *   - 遍历所有标题段落，关闭样式级自动编号，并按层级回填静态序号：
 *       二级=第X章（每遇一级重置）、三级=（N）、四级=N.、五级=（N，表格内，每遇上级重置）。
 *
 * 与 Python word_core_hyp1/hyp2.py 的 _strip_heading_auto_numbering / _find_heading1_id 对齐。
 */
public final class HeadingRenumberer {

    private HeadingRenumberer() {}

    // ===================== OOXML 命名空间 =====================
    private static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /** 中文数字（供「第X章」「（N）」及分类顶层「一、二、三」共用） */
    public static final String[] CHINESE_NUMS = {
        "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
        "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
        "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十"
    };

    // ===================== 对外 API =====================

    /** 查找（或动态注入）一级标题样式 ID */
    public static String findHeading1Id(XWPFDocument doc) {
        // 对齐 Python _find_heading1_id：按名称找 → 回退 ID 列表 → 注入 → 硬回退 '1'
        try {
            CTStyles styles = doc.getStyle();
            if (styles != null && styles.getStyleList() != null) {
                // Step 1：遍历全部样式，按名称（包含匹配，兼容"标题 1 自定义"等变体）
                for (CTStyle s : styles.getStyleList()) {
                    String n = styleName(s);
                    if (n == null) continue;
                    String lower = n.trim().toLowerCase();
                    if (lower.contains("标题 1") || lower.contains("heading 1") || lower.contains("heading1")) {
                        // 归一化：英文 heading 1 变体重命名为中文"标题 1"，
                        // 使样式栏显示与预期一致（仅改名称，styleId 不变 → 引用安全）
                        if (lower.contains("heading")) {
                            try {
                                CTString nm = s.getName();
                                if (nm != null) nm.setVal("\u6807\u9898 1"); // 标题 1
                            } catch (Exception ignore) {}
                        }
                        return s.getStyleId() != null ? s.getStyleId() : "1";
                    }
                }
                // Step 2：回退到常见 Heading 1 的 styleId
                for (String fid : new String[]{"1", "10", "11", "Heading1", "heading1"}) {
                    for (CTStyle s : styles.getStyleList()) {
                        if (fid.equals(s.getStyleId())) {
                            return fid;
                        }
                    }
                }
            }
        } catch (Exception ignore) {}

        // Step 3：动态注入（仅当基础文档完全无标题样式时尝试；lite 下可能失败）
        try {
            String injected = injectHeading1Style(doc);
            if (injected != null) return injected;
        } catch (Exception ignore) {}

        // Step 4：最终硬回退（对齐 Python return '1'）——保证 pStyle 永不缺失
        return "1";
    }

    /** 在 dst 文档插入一级分类标题（宋体 + 四号(14pt) + 加粗，无自动编号） */
    public static void insertCategoryHeading(XWPFDocument dst, String title, String h1Id) {
        XWPFParagraph p = dst.createParagraph();
        p.setAlignment(ParagraphAlignment.LEFT);
        // 注意：setAlignment 已（在 POI 内部）预建了第 1 个 pPr（含 w:jc）。
        // 必须复用该 pPr，绝不能再 addNewPPr()，否则会产生第 2 个 pPr，
        // 而 OOXML 规范下 Word/WPS 只认第一个 pPr，导致样式/编号/outlineLvl 全部失效。
        CTP ctp = p.getCTP();
        CTPPr pPr = ctp.getPPr();
        if (pPr == null) pPr = ctp.addNewPPr();

        // ── ① pStyle（schema 顺序：最前；样式栏显示"标题 1"的关键）──
        if (h1Id != null && !h1Id.isEmpty()) {
            pPr.addNewPStyle().setVal(h1Id);
        }

        // ── ② 禁用自动编号（numId=0 表示不参与任何列表）──
        CTNumPr numPr = pPr.addNewNumPr();
        numPr.addNewIlvl().setVal(BigInteger.ZERO);
        numPr.addNewNumId().setVal(BigInteger.ZERO);

        // ── ③ 强制 outlineLvl=0（schema 顺序：最后；导航窗格识别为一级标题）──
        CTDecimalNumber outlineLvl = pPr.addNewOutlineLvl();
        outlineLvl.setVal(BigInteger.ZERO);

        // ── ④ run 级格式（用户要求：宋体 + 四号(14pt) + 加粗）──
        XWPFRun r = p.createRun();
        r.setBold(true);
        r.setFontFamily("\u5B8B\u4F53"); // 宋体
        // 实测 POI 5.2.5 的 setFontSize(int) 以「磅(pt)」为单位（写入 sz 时会 ×2 成半磅）。
        // 四号 = 14pt，故直接传 14 → 落盘为 w:sz="28"（14pt），而非此前误传 28 导致的 w:sz="56"（28pt，大一倍）。
        r.setFontSize(14); // 一级标题字号 = 14pt（四号）
        r.setText(title);
    }

    /** 主处理：迭代所有 w:p 标题，禁用自动编号并回填静态序号（与 Python _strip_heading_auto_numbering 一致） */
    public static void stripHeadingAutoNumbering(XWPFDocument dst) {
        // 收集文档中存在的标题样式 ID → level（扫描全部样式，而非仅候选列表）
        Map<String, Integer> headingStyleIds = new HashMap<>();
        try {
            CTStyles styles = dst.getStyle();
            if (styles != null && styles.getStyleList() != null) {
                for (CTStyle s : styles.getStyleList()) {
                    int lvl = headingLevel(styleName(s));
                    if (lvl > 0 && s.getStyleId() != null) headingStyleIds.put(s.getStyleId(), lvl);
                }
            }
        } catch (Exception ignore) {}
        if (headingStyleIds.isEmpty()) {
            headingStyleIds.put("1", 1); headingStyleIds.put("2", 2);
            headingStyleIds.put("3", 3); headingStyleIds.put("4", 4); headingStyleIds.put("5", 5);
        }

        // 计数器（与 Python 一致的静态序号回填规则）
        //  H2(二级)=第X章，每遇 H1 重置；H3(三级)=（N），每遇 H2 重置；
        //  H4(四级)=N.，每遇 H3 重置；H5(五级)=（N），每遇 H1/H2/H3/H4 重置。
        int[] c2 = {0}, c3 = {0}, c4 = {0}, c5 = {0};

        Node bodyNode = dst.getDocument().getBody().getDomNode();
        walkAndAct(bodyNode, W_NS, "p", el -> {
            Element pPr = childEl(el, W_NS, "pPr");
            if (pPr == null) return;
            // 解析标题级别：优先按 pStyle 匹配；兜底——本程序插入的分类标题可能
            // 使用未登记到 headingStyleIds 的 styleId（findHeading1Id 的回退值），
            // 但统一带 outlineLvl=0，仍应识别为一级标题并触发计数器重置，
            // 否则二级标题第X章会跨分类持续递增（已确认 bug）。
            Element ps = childEl(pPr, W_NS, "pStyle");
            int lvl = 0;
            if (ps != null) {
                String sid = ps.getAttributeNS(W_NS, "val");
                if (sid != null && headingStyleIds.containsKey(sid)) lvl = headingStyleIds.get(sid);
            }
            if (lvl == 0) {
                Element outline = childEl(pPr, W_NS, "outlineLvl");
                if (outline != null && "0".equals(outline.getAttributeNS(W_NS, "val"))) lvl = 1;
            }
            if (lvl == 0) return;

            boolean wasAuto = hadAutoNumbering(pPr);

            // ① 所有标题：禁用自动编号（numPr 重定位到 pStyle 之后、numId=0）
            disableParaNumbering(pPr, el);

            // 表格内 且 原本非自动编号 且 非 H5 → 仅关闭自动编号，保持源文字原样
            // （注：表格内 H5 无论原先是否自动编号，均强制回填静态（N）以满足需求）
            if (isInsideTable(el) && !wasAuto && lvl != 5) return;

            String full = getFullText(el).trim();
            if (full.isEmpty()) return;

            // ---- 静态序号回填（先剥离旧编号前缀，再拼新前缀，避免叠加）----
            if (lvl == 1) {
                c2[0] = 0; c3[0] = 0; c4[0] = 0; c5[0] = 0;
            }
            else if (lvl == 2) {
                c2[0]++; c3[0] = 0; c4[0] = 0; c5[0] = 0;
                setTextPrefix(el, "第" + cn(c2[0]) + "章 ");
            }
            else if (lvl == 3) {
                c3[0]++; c4[0] = 0; c5[0] = 0;
                setTextPrefix(el, "（" + cn(c3[0]) + "）");
            }
            else if (lvl == 4) {
                c4[0]++; c5[0] = 0;
                setTextPrefix(el, c4[0] + ". ");
            }
            else if (lvl == 5) {
                // H5（表格内）回填静态「（N）」，N 用阿拉伯数字；每遇 H1/H2/H3/H4 已归零
                c5[0]++;
                setTextPrefix(el, "（" + c5[0] + "）");
            }
        });
    }

    /** 安全读取 CTStyle 的 w:name/@w:val（poi-ooxml-lite 下 XWPFStyle 名取自 CTStyle）；供 WordMergeService.buildStyleMap 复用 */
    public static String styleName(CTStyle s) {
        if (s == null) return null;
        CTString nm = s.getName();
        return nm != null ? nm.getVal() : null;
    }

    // ===================== 私有实现 =====================

    /** 段落直接格式是否含有效 numPr/numId (>0)，即曾参与 Word 自动编号 */
    private static boolean hadAutoNumbering(Element pPr) {
        Element numPr = childEl(pPr, W_NS, "numPr");
        if (numPr == null) return false;
        Element numId = childEl(numPr, W_NS, "numId");
        if (numId == null) return false;
        String val = numId.getAttributeNS(W_NS, "val");
        return val != null && !"0".equals(val) && !val.isEmpty();
    }

    /** 禁用段落样式级自动编号：numPr 位于 pStyle 之后（OOXML 顺序），numId=0 */
    private static void disableParaNumbering(Element pPr, Element pEl) {
        Element pStyle = childEl(pPr, W_NS, "pStyle");
        Element numPr = childEl(pPr, W_NS, "numPr");
        if (numPr == null) {
            numPr = pEl.getOwnerDocument().createElementNS(W_NS, "w:numPr");
            if (pStyle != null) {
                Node next = pStyle.getNextSibling();
                if (next != null) pPr.insertBefore(numPr, next);
                else pPr.appendChild(numPr);
            } else {
                pPr.insertBefore(numPr, pPr.getFirstChild());
            }
        } else if (pStyle != null) {
            // 已存在但可能位于 pPr 末尾（错误位置）→ 重定位到 pStyle 之后，
            // 否则 Word 可能忽略该 numPr，导致样式级自动编号仍生效（跨段延续）。
            Node next = pStyle.getNextSibling();
            if (next != null) {
                if (next != numPr) pPr.insertBefore(numPr, next);
            } else {
                pPr.appendChild(numPr);
            }
        }
        Element numId = childEl(numPr, W_NS, "numId");
        if (numId == null) {
            numId = pEl.getOwnerDocument().createElementNS(W_NS, "w:numId");
            numPr.appendChild(numId);
        }
        numId.setAttributeNS(W_NS, "w:val", "0");
        Element ilvl = childEl(numPr, W_NS, "ilvl");
        if (ilvl == null) {
            ilvl = pEl.getOwnerDocument().createElementNS(W_NS, "w:ilvl");
            numPr.insertBefore(ilvl, numId);
        }
        ilvl.setAttributeNS(W_NS, "w:val", "0");
    }

    /** 段落是否位于表格（w:tbl）内 */
    private static boolean isInsideTable(Node pNode) {
        Node cur = pNode.getParentNode();
        while (cur != null) {
            if (cur.getNodeType() == Node.ELEMENT_NODE
                    && "tbl".equals(cur.getLocalName())
                    && W_NS.equals(cur.getNamespaceURI())) {
                return true;
            }
            cur = cur.getParentNode();
        }
        return false;
    }

    private static int headingLevel(String name) {
        if (name == null) return 0;
        // 源文档样式名多为小写 heading N（如 "heading 1"）；Python 基线用 re.IGNORECASE，
        // 故此处必须忽略大小写，否则级别判定失败导致全部标题不被重编号。
        Matcher m = Pattern.compile("^(heading|标题)\\s*(\\d+)$", Pattern.CASE_INSENSITIVE).matcher(name.trim());
        if (!m.find()) return 0;
        try {
            int l = Integer.parseInt(m.group(2));
            return (l >= 1 && l <= 5) ? l : 0;
        } catch (Exception e) { return 0; }
    }

    private static String cn(int n) {
        return (n >= 1 && n <= CHINESE_NUMS.length) ? CHINESE_NUMS[n - 1] : String.valueOf(n);
    }

    /**
     * 向 dst 文档的 styles.xml 动态注入「标题 1 / Heading 1」段落样式。
     * 使用 xmlbeans 原生 CT 对象 API（非 DOM），确保 poi-ooxml-lite 序列化时样式定义被正确写入 .docx。
     */
    private static String injectHeading1Style(XWPFDocument doc) {
        try {
            CTStyles ct = doc.getStyle();
            if (ct == null) return null;

            // ── 生成不冲突的 styleId ──
            String newId = "Heading1Merged";
            List<CTStyle> existing = ct.getStyleList();
            if (existing != null) {
                int suffix = 1;
                boolean collision;
                do {
                    collision = false;
                    for (CTStyle s : existing) {
                        if (s.getStyleId() != null && newId.equals(s.getStyleId())) {
                            newId = "Heading1Merged" + (suffix++);
                            collision = true;
                            break;
                        }
                    }
                } while (collision);
            }

            // ── 用 xmlbeans 原生 API 构建完整样式定义（序列化安全）──
            CTStyle newStyle = ct.addNewStyle();
            newStyle.setStyleId(newId);
            newStyle.setType(STStyleType.PARAGRAPH);

            // 名称
            newStyle.addNewName().setVal("\u6807\u9898 1"); // "标题 1"

            // basedOn → Normal / 正文
            String normalId = findNormalStyleId(doc);
            if (normalId != null) {
                newStyle.addNewBasedOn().setVal(normalId);
            }

            // ── 段落属性 pPr ──
            CTPPrGeneral pPr = newStyle.addNewPPr();

            // outlineLvl=0 → 一级标题（导航窗格识别关键）
            pPr.addNewOutlineLvl().setVal(BigInteger.ZERO);

            // 段前/段后间距
            CTSpacing spacing = pPr.addNewSpacing();
            spacing.setBefore(BigInteger.valueOf(240));
            spacing.setAfter(BigInteger.valueOf(120));

            // 左对齐
            pPr.addNewJc().setVal(STJc.LEFT);

            // ── 字符属性 rPr（黑体 + 粗体 + 16pt）──
            CTRPr rPr = newStyle.addNewRPr();
            CTFonts fonts = rPr.addNewRFonts();
            fonts.setAscii("\u9ED1\u4F53");     // "黑体"
            fonts.setEastAsia("\u9ED1\u4F53");   // "黑体"
            rPr.addNewB().setVal(true);
            rPr.addNewSz().setVal("32");          // 16pt（半磅单位）
            rPr.addNewSzCs().setVal("32");

            return newId;
        } catch (Exception e) {
            return null;
        }
    }

    /** 查找 Normal / 正文 样式的真实 ID（作为注入样式的 basedOn 基准） */
    private static String findNormalStyleId(XWPFDocument doc) {
        try {
            CTStyles styles = doc.getStyle();
            if (styles == null || styles.getStyleList() == null) return null;
            for (CTStyle s : styles.getStyleList()) {
                String nm = styleName(s);
                if (nm == null) continue;
                nm = nm.trim();
                if ("Normal".equals(nm) || "正文".equals(nm)) {
                    return s.getStyleId() != null ? s.getStyleId() : "Normal";
                }
            }
        } catch (Exception ignore) {}
        return null;
    }
}
