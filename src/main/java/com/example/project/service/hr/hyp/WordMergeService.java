package com.example.project.service.hr.hyp;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.math.BigInteger;
import java.security.MessageDigest;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.example.project.service.hr.hyp.OoxmlDomUtils.*;

/**
 * 对应 Python: word_core_hyp1.py + word_core_hyp2.py（方案A：POI XWPF + OOXML XML 层）
 * 功能：Word 文档「分类层级合并」。
 *   1) 类别关键词提取（0-X分 模式 + 主观回退，带 parent 类别名）
 *   2) 智能搜索（精准 → 模糊 4字分块两级）
 *   3) 合并引擎（CT XML 深克隆 + 图片 MD5 去重重映射 + 样式映射 + 编号合并 + A4 分节符 + 标题静态编号回填）
 *
 * 与 Python 的一致性要点（相对上一版的修正）：
 *   - detectParentCol：取「最靠左、文字最短」的候选（上一版取了最长，类别名选错）
 *   - 关键词清洗：先剥尾部标点（含全角句号 。），再剥尾部括号说明（锚定 $），与 Python 一致
 *   - copyAllContent：把【源样式级】numPr 提升到段落直接格式（_promote_para），
 *     使 _strip_heading_auto_numbering 能正确区分「原为自动编号」与「原为纯文本」标题
 *   - mergeNumbering：真正合并 numbering.xml（abstractNum + num，冲突重分配 ID），而非空实现
 *   - stripHeadingAutoNumbering：H5 同样回填静态序号「（N）」（阿拉伯数字），并扫描全部样式识别标题
 */
@Service
public class WordMergeService {

    // ===================== OOXML 命名空间 =====================
    private static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final String R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static final String A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String V_NS = "urn:schemas-microsoft-com:vml";
    private static final String O_NS = "urn:schemas-microsoft-com:office:office";

    // ===================== 常量 =====================
    private static final long TWIPS_PER_CM = 567;
    private static final long A4_W = Math.round(21.0 * TWIPS_PER_CM);   // 11907
    private static final long A4_H = Math.round(29.7 * TWIPS_PER_CM);   // 16840
    private static final long MARGIN_2CM = 2 * TWIPS_PER_CM;   // 1134
    private static final long MARGIN_18CM = Math.round(1.8 * TWIPS_PER_CM); // 1021
    private static final long MARGIN_15CM = Math.round(1.5 * TWIPS_PER_CM); // 850

    /** 中英文标题样式别名映射（处理中英文 Word 差异） */
    private static final Map<String, List<String>> BUILTIN_MAP = new LinkedHashMap<>();

    static {
        String[][] m = {
            {"标题 1", "Heading 1", "标题1"}, {"标题 2", "Heading 2", "标题2"},
            {"标题 3", "Heading 3", "标题3"}, {"标题 4", "Heading 4", "标题4"},
            {"标题 5", "Heading 5", "标题5"}, {"标题 6", "Heading 6", "标题6"},
            {"标题 7", "Heading 7", "标题7"}, {"标题 8", "Heading 8", "标题8"},
            {"标题 9", "Heading 9", "标题9"}
        };
        for (String[] row : m) {
            for (String s : row) BUILTIN_MAP.put(s, Arrays.asList(row));
        }
    }

    // ===================== DTO =====================
    /** 匹配方式常量 */
    public static final String MATCH_PRECISE = "precise";   // 精准匹配（文件名含完整关键词）
    public static final String MATCH_FUZZY = "fuzzy";       // 模糊匹配（按4字分块命中）

    /** 单个匹配文件及其匹配方式 */
    public static class MatchedFile {
        public String fileName;
        public String matchType;   // MATCH_PRECISE / MATCH_FUZZY

        public MatchedFile() {}
        public MatchedFile(String fileName, String matchType) {
            this.fileName = fileName;
            this.matchType = matchType;
        }

        @Override
        public String toString() {
            return fileName + (MATCH_FUZZY.equals(matchType) ? "（模糊）" : "");
        }
    }

    public static class Category {
        public String parent;
        public String type;
        public List<String> keywords = new ArrayList<>();
        public List<MatchedFile> matchedFiles = new ArrayList<>();
    }

    // ===================== 校验 =====================
    /**
     * 安全打开 Word 文档：先校验文件头，再加载。
     * 拒绝旧版 .doc（OLE2: D0CF11E0）和无效 ZIP。
     */
    public XWPFDocument openValidatedDoc(MultipartFile mf) throws Exception {
        byte[] head = new byte[4];
        try (InputStream is = mf.getInputStream()) {
            int n = is.read(head);
            if (n >= 2 && head[0] == 'P' && head[1] == 'K') {
                // docx (ZIP) —— 继续
            } else if (n >= 4 && (head[0] & 0xFF) == 0xD0 && (head[1] & 0xFF) == 0xCF
                    && (head[2] & 0xFF) == 0x11 && (head[3] & 0xFF) == 0xE0) {
                throw new IllegalArgumentException("文件「" + mf.getOriginalFilename()
                        + "」是旧版 .doc 格式（Word 97-2003），请另存为 .docx 后重试");
            } else {
                throw new IllegalArgumentException("文件「" + mf.getOriginalFilename() + "」不是有效的 .docx 文件");
            }
        }
        try (InputStream is = mf.getInputStream()) {
            return new XWPFDocument(is);
        }
    }

    // ===================== 类别提取 =====================
    public List<Category> extractCategories(List<MultipartFile> standards) throws Exception {
        Map<String, Category> merged = new LinkedHashMap<>();
        for (MultipartFile mf : standards) {
            try (XWPFDocument doc = openValidatedDoc(mf)) {
                List<Category> cats = new ArrayList<>();
                cats.addAll(CategoryExtractor.extractScoreCategories(doc));        // 一级：0-X分
                cats.addAll(CategoryExtractor.extractSubjectiveCategories(doc));   // 二级：主观回退
                cats.addAll(CategoryExtractor.extractMaxscoreCategories(doc, null)); // 三级：满分值模式
                for (Category c : cats) {
                    Category ex = merged.get(c.parent);
                    if (ex == null) {
                        ex = new Category();
                        ex.parent = c.parent;
                        ex.type = c.type;
                        merged.put(c.parent, ex);
                    }
                    for (String k : c.keywords) {
                        if (!ex.keywords.contains(k)) ex.keywords.add(k);
                    }
                }
            }
        }
        if (merged.isEmpty()) {
            throw new IllegalArgumentException(
                    "文档中未找到包含\"0-X分\"、\"主观\"或\"满分值\"评分标准的行，无法提取关键词");
        }
        return new ArrayList<>(merged.values());
    }


    // ===================== 搜索 =====================
    public List<Category> searchCategories(List<Category> categories, List<String> names) {
        List<String> all = new ArrayList<>(names);
        Collections.sort(all);
        for (Category c : categories) {
            // 文件名 → 匹配方式（精准优先：已被精准命中者不再降级为模糊）
            Map<String, String> matchMap = new LinkedHashMap<>();
            for (String kw : c.keywords) {
                List<String> precise = new ArrayList<>();
                for (String f : all) {
                    if (f.toLowerCase().contains(kw.toLowerCase())) precise.add(f);
                }
                if (!precise.isEmpty()) {
                    for (String f : precise) {
                        // 精准命中（若此前已是精准则保持不变）
                        matchMap.putIfAbsent(f, MATCH_PRECISE);
                    }
                    continue;
                }
                for (String fw : CategoryExtractor.fuzzyChunks(kw, 4)) {
                    for (String f : all) {
                        if (f.toLowerCase().contains(fw.toLowerCase())) {
                            // 仅当尚未被精准命中时记为模糊
                            matchMap.putIfAbsent(f, MATCH_FUZZY);
                        }
                    }
                }
            }
            List<MatchedFile> matched = new ArrayList<>();
            for (Map.Entry<String, String> e : matchMap.entrySet()) {
                matched.add(new MatchedFile(e.getKey(), e.getValue()));
            }
            c.matchedFiles = matched;
        }
        return categories;
    }

    // ===================== 合并引擎 =====================
    public String mergeDocuments(List<Category> categories, Map<String, MultipartFile> candidateMap,
                                 String outputDir, List<String> logs) throws Exception {
        if (logs == null) logs = new ArrayList<>();
        // 全局去重：同一文件只保留在首个出现的分类中
        LinkedHashSet<String> seen = new LinkedHashSet<>();
        List<String> allFiles = new ArrayList<>();
        for (Category c : categories) {
            for (MatchedFile mf : c.matchedFiles) {
                if (seen.add(mf.fileName.toLowerCase())) allFiles.add(mf.fileName);
            }
        }
        logLine(logs, "▶ 开始分类合并：共 " + categories.size() + " 个分类、"
                + candidateMap.size() + " 个候选文件，去重后待合并 " + allFiles.size() + " 个");
        if (allFiles.isEmpty()) {
            logLine(logs, "✗ 没有匹配到任何文件，合并终止");
            throw new IllegalArgumentException("没有匹配到任何文件");
        }

        MultipartFile baseFile = candidateMap.get(allFiles.get(0));
        if (baseFile == null) {
            logLine(logs, "✗ 找不到基础文件：" + allFiles.get(0) + "，合并终止");
            throw new IllegalArgumentException("找不到基础文件：" + allFiles.get(0));
        }

        XWPFDocument dst = openValidatedDoc(baseFile);
        // 统一基础文档页边距
        CTSectPr baseSect = dst.getDocument().getBody().getSectPr();
        if (baseSect != null) setSectPrMargins(baseSect);
        // 提取基础文档的页眉/页脚 rId（在所有分节符中沿用）
        Map<String, String> hfRids = extractBaseHfRids(dst);
        // 清空 body 内容（保留 sectPr）
        clearBodyKeepSectPr(dst);
        logLine(logs, "✓ 基础文档（样式/页眉页脚基准）：" + baseFile.getOriginalFilename());
        logLine(logs, "✓ 已清空基础文档正文，准备写入分类层级");

        String h1Id = HeadingRenumberer.findHeading1Id(dst);
        Map<String, String> md5Cache = new HashMap<>();

        int catIndex = 0;
        Set<String> processed = new HashSet<>();
        for (Category cat : categories) {
            List<String> files = new ArrayList<>();
            for (MatchedFile mf : cat.matchedFiles) {
                if (processed.contains(mf.fileName.toLowerCase())) continue;
                if (!candidateMap.containsKey(mf.fileName)) continue;
                files.add(mf.fileName);
            }
            if (files.isEmpty()) continue;

            catIndex++;
            String num = catIndex <= HeadingRenumberer.CHINESE_NUMS.length ? HeadingRenumberer.CHINESE_NUMS[catIndex - 1] : String.valueOf(catIndex);
            String heading = num + "、" + cat.parent;

            if (catIndex > 1) insertSectionBreak(dst, hfRids);
            HeadingRenumberer.insertCategoryHeading(dst, heading, h1Id);
            logLine(logs, "── 分类 " + catIndex + "/" + categories.size() + "：" + heading
                    + "（命中 " + files.size() + " 个文件）──");

            for (int fi = 0; fi < files.size(); fi++) {
                String fname = files.get(fi);
                MultipartFile mf = candidateMap.get(fname);
                long kb = mf != null ? mf.getSize() / 1024 : 0;
                try (XWPFDocument src = openValidatedDoc(mf)) {
                    // 1) 合并 numbering（必须在 copy 之前，返回 numId 重映射表）
                    Map<String, String> numIdMap = mergeNumbering(src, dst);
                    // 2) 构建样式映射
                    Map<String, String> styleMap = buildStyleMap(src, dst);
                    // 3) 文件间插入分节符
                    if (fi > 0) insertSectionBreak(dst, hfRids);
                    // 4) 复制内容（含 _promote_para + 图片重映射 + 样式/编号映射）
                    copyAllContent(src, dst, md5Cache, styleMap, numIdMap);
                    processed.add(fname.toLowerCase());
                    logLine(logs, "   " + (fi + 1) + ". 追加：" + fname + "（" + kb + " KB）");
                } catch (Exception e) {
                    logLine(logs, "   ⚠ 跳过失败文档：" + fname + " —— " + e.getMessage());
                }
            }
        }

        // 标题静态编号回填（关闭样式级自动编号，并回填静态文字前缀）
        logLine(logs, "✓ 内容复制完成，开始标题静态编号回填（关闭样式级自动编号 + 回填静态“第X章/（一）/1./（1）”等）");
        HeadingRenumberer.stripHeadingAutoNumbering(dst);
        logLine(logs, "✓ 标题编号回填完成");

        File dir = new File(outputDir);
        if (!dir.exists()) dir.mkdirs();
        // 使用带时间戳的唯一文件名，避免并发合并互相覆盖
        String outName = "文档合并结果_" + System.currentTimeMillis() + ".docx";
        File out = new File(dir, outName);
        try (FileOutputStream fos = new FileOutputStream(out)) {
            dst.write(fos);
        }
        dst.close();
        double mb = out.length() / 1048576.0;
        logLine(logs, "▶ 写出文件：" + outName + "（" + String.format("%.2f", mb) + " MB）");
        logLine(logs, "■ 合并结束");
        return out.getAbsolutePath();
    }

    private static final DateTimeFormatter HHMMSS = DateTimeFormatter.ofPattern("HH:mm:ss");
    private void logLine(List<String> logs, String msg) {
        if (logs != null) logs.add(LocalTime.now().format(HHMMSS) + "  " + msg);
    }

    // ---------- 清空 body（保留 sectPr） ----------
    private void clearBodyKeepSectPr(XWPFDocument doc) {
        CTBody body = doc.getDocument().getBody();
        for (int i = body.getPList().size() - 1; i >= 0; i--) {
            body.removeP(i);
        }
        for (int i = body.getTblList().size() - 1; i >= 0; i--) {
            body.removeTbl(i);
        }
    }

    // ---------- 分节符（下一页 + A4 + 统一页边距 + 页眉页脚沿用） ----------
    private void insertSectionBreak(XWPFDocument dst, Map<String, String> hfRids) {
        XWPFParagraph p = dst.createParagraph();
        CTP ctp = p.getCTP();
        CTPPr pPr = ctp.addNewPPr();
        CTSectPr sect = pPr.addNewSectPr();
        sect.addNewType().setVal(STSectionMark.NEXT_PAGE);
        setSectPrMargins(sect);
        if (hfRids != null && !hfRids.isEmpty()) {
            try {
                Node sectNode = sect.getDomNode();
                org.w3c.dom.Document domDoc = sectNode.getOwnerDocument();
                for (Map.Entry<String, String> e : hfRids.entrySet()) {
                    String[] parts = e.getKey().split(":", 2);
                    String refName = parts[0].equals("header") ? "w:headerReference" : "w:footerReference";
                    String type = parts[1];
                    Element ref = domDoc.createElementNS(W_NS, refName);
                    ref.setAttributeNS(W_NS, "w:type", type);
                    ref.setAttributeNS(R_NS, "r:id", e.getValue());
                    sectNode.appendChild(ref);
                }
            } catch (Exception ignore) {}
        }
    }

    /** 从基础文档 sectPr 中提取 headerReference / footerReference 的 rId 映射 */
    private Map<String, String> extractBaseHfRids(XWPFDocument doc) {
        Map<String, String> hfRids = new HashMap<>();
        try {
            CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
            if (sectPr == null) return hfRids;
            Node sectNode = sectPr.getDomNode();
            NodeList children = sectNode.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node child = children.item(i);
                if (child.getNodeType() != Node.ELEMENT_NODE) continue;
                Element el = (Element) child;
                String local = el.getLocalName();
                if ("headerReference".equals(local)) {
                    String type = el.getAttributeNS(W_NS, "type");
                    String rid = el.getAttributeNS(R_NS, "id");
                    if (type == null) type = "default";
                    if (rid != null) hfRids.put("header:" + type, rid);
                } else if ("footerReference".equals(local)) {
                    String type = el.getAttributeNS(W_NS, "type");
                    String rid = el.getAttributeNS(R_NS, "id");
                    if (type == null) type = "default";
                    if (rid != null) hfRids.put("footer:" + type, rid);
                }
            }
        } catch (Exception ignore) {}
        return hfRids;
    }

    /** 统一设置分节符的纸张与页边距（DOM 直接写 w:pgSz / w:pgMar） */
    private void setSectPrMargins(CTSectPr sect) {
        try {
            org.w3c.dom.Node sectNode = sect.getDomNode();
            org.w3c.dom.Document domDoc = sectNode.getOwnerDocument();
            org.w3c.dom.NodeList children = sectNode.getChildNodes();
            for (int i = children.getLength() - 1; i >= 0; i--) {
                org.w3c.dom.Node child = children.item(i);
                String localName = child.getLocalName();
                if ("pgSz".equals(localName) || "pgMar".equals(localName)) {
                    sectNode.removeChild(child);
                }
            }
            org.w3c.dom.Element pgSz = domDoc.createElementNS(W_NS, "w:pgSz");
            pgSz.setAttributeNS(W_NS, "w:w", String.valueOf(A4_W));
            pgSz.setAttributeNS(W_NS, "w:h", String.valueOf(A4_H));
            pgSz.setAttributeNS(W_NS, "w:orient", "portrait");
            sectNode.appendChild(pgSz);
            org.w3c.dom.Element pgMar = domDoc.createElementNS(W_NS, "w:pgMar");
            pgMar.setAttributeNS(W_NS, "w:top", String.valueOf(MARGIN_2CM));
            pgMar.setAttributeNS(W_NS, "w:bottom", String.valueOf(MARGIN_2CM));
            pgMar.setAttributeNS(W_NS, "w:left", String.valueOf(MARGIN_18CM));
            pgMar.setAttributeNS(W_NS, "w:right", String.valueOf(MARGIN_18CM));
            pgMar.setAttributeNS(W_NS, "w:gutter", "0");
            pgMar.setAttributeNS(W_NS, "w:header", String.valueOf(MARGIN_2CM));
            pgMar.setAttributeNS(W_NS, "w:footer", String.valueOf(MARGIN_15CM));
            sectNode.appendChild(pgMar);
        } catch (Exception e) {
            // 静默失败，不影响主流程
        }
    }

    // ---------- 分类标题（Heading 1 + 禁用自动编号） ----------

    // ---------- 深克隆 body 内容 + 图片重映射 + 样式/编号映射 + 源样式级 numPr 提升 ----------
    private void copyAllContent(XWPFDocument src, XWPFDocument dst, Map<String, String> md5Cache,
                                Map<String, String> styleMap, Map<String, String> numIdMap) throws Exception {
        CTBody srcBody = src.getDocument().getBody();
        CTBody dstBody = dst.getDocument().getBody();

        // 源文档图片关系
        Map<String, XWPFPictureData> srcRelToPic = new HashMap<>();
        try {
            for (XWPFPictureData pic : src.getAllPictures()) {
                try { srcRelToPic.put(src.getRelationId(pic), pic); } catch (Exception ignore) {}
            }
        } catch (Exception ignore) {}

        // 顺序深克隆（跳过 sectPr），并同步记录 (源元素, 目标克隆) 配对
        List<Element> srcEls = new ArrayList<>();
        List<Element> dstClones = new ArrayList<>();
        Node srcBodyNode = srcBody.getDomNode();
        NodeList srcKids = srcBodyNode.getChildNodes();
        for (int i = 0; i < srcKids.getLength(); i++) {
            Node n = srcKids.item(i);
            if (n.getNodeType() != Node.ELEMENT_NODE) continue;
            Element e = (Element) n;
            if (!W_NS.equals(e.getNamespaceURI())) continue;
            String local = e.getLocalName();
            if ("p".equals(local) || "tbl".equals(local)) srcEls.add(e);
        }

        Node dstBodyNode = dstBody.getDomNode();
        org.w3c.dom.Document dstDoc = dstBodyNode.getOwnerDocument();
        // 文档级 sectPr 必须位于 body 末尾，内容需插在其之前（与 python-docx Body.append 一致）
        Node trailingSectPr = null;
        NodeList bk = dstBodyNode.getChildNodes();
        for (int i = bk.getLength() - 1; i >= 0; i--) {
            Node c = bk.item(i);
            if (c.getNodeType() == Node.ELEMENT_NODE && "sectPr".equals(c.getLocalName())
                    && W_NS.equals(c.getNamespaceURI())) {
                trailingSectPr = c;
                break;
            }
        }
        for (Element srcEl : srcEls) {
            // 复制前先记录源段落的 pStyle（用于源样式级 numPr 提升，必须在 remap 样式前取值）
            String srcStyleVal = getPStyleVal(srcEl);

            // 跨文档深克隆：importNode + insertBefore(sectPr)（xmlbeans 实时 DOM，随 commit 持久化）
            Node imported = dstDoc.importNode(srcEl, true);
            if (trailingSectPr != null) dstBodyNode.insertBefore(imported, trailingSectPr);
            else dstBodyNode.appendChild(imported);
            Element dstClone = (Element) imported;
            if (dstClone == null) continue;
            dstClones.add(dstClone);

            // 应用样式映射（方案 A：使用基础文档样式体系）
            if (styleMap != null && !styleMap.isEmpty()) applyStyleMapElement(dstClone, styleMap);

            // 把【源样式级】numPr 提升到段落直接格式
            if ("p".equals(srcEl.getLocalName())) {
                promotePara(src, dst, srcStyleVal, dstClone);
            } else {
                List<Element> srcPs = new ArrayList<>();
                collectAllElems(srcEl, "p", srcPs);
                List<Element> dstPs = new ArrayList<>();
                collectAllElems(dstClone, "p", dstPs);
                for (int k = 0; k < srcPs.size() && k < dstPs.size(); k++) {
                    promotePara(src, dst, getPStyleVal(srcPs.get(k)), dstPs.get(k));
                }
            }

            // 图片重映射（覆盖 dstClone 子树）
            remapImages(dst, srcRelToPic, md5Cache, dstClone);

            // 编号 ID 重映射（numbering.xml 合并后的 numId 对齐）
            if (numIdMap != null && !numIdMap.isEmpty()) applyNumIdMapElement(dstClone, numIdMap);
        }
    }

    /**
     * 把【源样式级】w:numPr 提升到目标段落直接格式（对应 Python _promote_para）。
     * 仅当源段落直接格式无 numPr、且源样式定义含有效 numId 时，将源样式里的 w:numPr
     * 深拷贝到目标段落 w:pPr（pStyle 之后）。
     * 目的：让后续 _strip_heading_auto_numbering 仅依据段落直接 numPr 判断标题是否
     * 原为自动编号，从而正确区分「原为自动编号」与「原为纯文本」的同类标题。
     */
    private void promotePara(XWPFDocument src, XWPFDocument dst, String srcStyleVal, Element dstP) {
        if (srcStyleVal == null || srcStyleVal.isEmpty()) return;
        // 从【源文档】样式定义读取 numId（源样式级 numPr → 段落直接格式）
        String numId = getStyleNumId(src, srcStyleVal);
        if (numId == null) return;

        Element dstPPr = childEl(dstP, W_NS, "pPr");
        if (dstPPr == null) return;
        // 目标段落直接格式已含 numPr（含 numId=0 关闭编号）则不覆盖
        if (childEl(dstPPr, W_NS, "numPr") != null) return;

        org.w3c.dom.Document doc = dstP.getOwnerDocument();
        Element numPr = doc.createElementNS(W_NS, "w:numPr");
        Element ilvl = doc.createElementNS(W_NS, "w:ilvl");
        ilvl.setAttributeNS(W_NS, "w:val", "0");
        Element numIdEl = doc.createElementNS(W_NS, "w:numId");
        numIdEl.setAttributeNS(W_NS, "w:val", numId);
        numPr.appendChild(ilvl);
        numPr.appendChild(numIdEl);
        Element pStyle = childEl(dstPPr, W_NS, "pStyle");
        if (pStyle != null) {
            Node next = pStyle.getNextSibling();
            if (next != null) dstPPr.insertBefore(numPr, next);
            else dstPPr.appendChild(numPr);
        } else {
            dstPPr.insertBefore(numPr, dstPPr.getFirstChild());
        }
    }

    /**
     * 读取源段落所用样式（styleId）的【样式级】numPr 中的 numId（>0 表示原为自动编号）。
     * 通过 XMLBeans 直接访问 CTStyle，避免依赖 getCTStyles()（poi-ooxml-lite 下不存在）。
     */
    private String getStyleNumId(XWPFDocument doc, String styleId) {
        try {
            XWPFStyle s = doc.getStyles().getStyle(styleId);
            if (s == null) return null;
            CTStyle ct = s.getCTStyle();
            CTPPrGeneral pPr = ct.getPPr();
            if (pPr == null) return null;
            CTNumPr numPr = pPr.getNumPr();
            if (numPr == null) return null;
            CTDecimalNumber numId = numPr.getNumId();
            if (numId == null || numId.getVal() == null) return null;
            String v = numId.getVal().toString();
            if ("0".equals(v)) return null;
            return v;
        } catch (Exception e) {
            return null;
        }
    }

    // ---------- 图片重映射（MD5 去重） ----------
    private void remapImages(XWPFDocument dst, Map<String, XWPFPictureData> srcRelToPic,
                             Map<String, String> md5Cache, Node subtree) {
        walkAndAct(subtree, A_NS, "blip", el -> {
            String rid = el.getAttributeNS(R_NS, "embed");
            if (rid == null || rid.isEmpty()) rid = el.getAttributeNS(R_NS, "link");
            if (rid == null || rid.isEmpty()) return;
            XWPFPictureData pic = srcRelToPic.get(rid);
            if (pic == null) return;
            try {
                String newRid = ensurePictureInDst(dst, pic, md5Cache);
                el.setAttributeNS(R_NS, "r:embed", newRid);
                el.removeAttributeNS(R_NS, "link");
            } catch (Exception ignore) {}
        });

        walkAndAct(subtree, V_NS, "imagedata", el -> {
            String rid = el.getAttributeNS(R_NS, "id");
            if (rid == null || rid.isEmpty()) rid = el.getAttributeNS(O_NS, "relid");
            if (rid == null || rid.isEmpty()) return;
            XWPFPictureData pic = srcRelToPic.get(rid);
            if (pic == null) return;
            try {
                String newRid = ensurePictureInDst(dst, pic, md5Cache);
                if (el.hasAttributeNS(R_NS, "id")) el.setAttributeNS(R_NS, "r:id", newRid);
                else if (el.hasAttributeNS(O_NS, "relid")) el.setAttributeNS(O_NS, "o:relid", newRid);
            } catch (Exception ignore) {}
        });
    }

    private String ensurePictureInDst(XWPFDocument dst, XWPFPictureData pic, Map<String, String> md5Cache) throws Exception {
        byte[] data = pic.getData();
        String md5 = md5Hex(data);
        if (md5Cache.containsKey(md5)) return md5Cache.get(md5);
        int format = pic.getPictureType();
        String rid;
        try (java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream(data)) {
            rid = dst.addPictureData(bais, format);
        }
        md5Cache.put(md5, rid);
        return rid;
    }

    // ---------- 样式映射（DOM 操作子树） ----------
    private void applyStyleMapElement(Node subtree, Map<String, String> styleMap) {
        walkAndAct(subtree, W_NS, "pStyle", el -> {
            String v = el.getAttributeNS(W_NS, "val");
            if (v != null && styleMap.containsKey(v)) el.setAttributeNS(W_NS, "val", styleMap.get(v));
        });
        walkAndAct(subtree, W_NS, "tblStyle", el -> {
            String v = el.getAttributeNS(W_NS, "val");
            if (v != null && styleMap.containsKey(v)) el.setAttributeNS(W_NS, "val", styleMap.get(v));
        });
    }

    private void applyNumIdMapElement(Node subtree, Map<String, String> numIdMap) {
        walkAndAct(subtree, W_NS, "numId", el -> {
            String v = el.getAttributeNS(W_NS, "val");
            if (v != null && numIdMap.containsKey(v)) el.setAttributeNS(W_NS, "val", numIdMap.get(v));
        });
    }

    /**
     * 构建源→目标样式 ID 映射。扫描双文档全部样式，按名称匹配（含中英文别名）。
     */
    private Map<String, String> buildStyleMap(XWPFDocument src, XWPFDocument dst) {
        Map<String, String> map = new HashMap<>();
        try {
            // poi-ooxml-lite 下 XWPFStyles 无 getStyleList()，改用实时 CTStyles.getStyleList()
            CTStyles srcCt = src.getStyle();
            CTStyles dstCt = dst.getStyle();
            if (srcCt == null || dstCt == null) return map;
            List<CTStyle> dstList = dstCt.getStyleList();
            List<CTStyle> srcList = srcCt.getStyleList();
            if (dstList == null || srcList == null) return map;

            Map<String, String> dstNameToId = new HashMap<>();
            for (CTStyle s : dstList) {
                String name = HeadingRenumberer.styleName(s);
                String id = s.getStyleId();
                if (name != null && id != null) {
                    dstNameToId.putIfAbsent(name.trim(), id);
                }
            }

            for (CTStyle s : srcList) {
                String srcName = HeadingRenumberer.styleName(s);
                if (srcName == null) continue;
                srcName = srcName.trim();
                String srcId = s.getStyleId();
                if (srcId == null) continue;

                String dstId = dstNameToId.get(srcName);
                if (dstId != null && !dstId.equals(srcId)) {
                    map.put(srcId, dstId);
                    continue;
                }
                dstId = matchBuiltinAlias(srcName, dstNameToId);
                if (dstId != null && !dstId.equals(srcId)) {
                    map.put(srcId, dstId);
                }
            }
        } catch (Exception ignore) {}
        return map;
    }

    private String matchBuiltinAlias(String srcName, Map<String, String> dstNameToId) {
        for (Map.Entry<String, List<String>> e : BUILTIN_MAP.entrySet()) {
            if (srcName.equals(e.getKey()) || e.getValue().contains(srcName)) {
                if (dstNameToId.containsKey(e.getKey())) return dstNameToId.get(e.getKey());
                for (String alias : e.getValue()) {
                    if (dstNameToId.containsKey(alias)) return dstNameToId.get(alias);
                }
            }
        }
        return null;
    }

    // ---------- 编号合并（真正合并 numbering.xml，基于 POI XMLBeans CT API） ----------
    /**
     * 将源文档 numbering.xml 中缺失的列表格式定义复制到目标文档。
     * 与 Python _merge_numbering 一致：
     *   1) 遍历源 abstractNum，目标缺失的直接复制；冲突（ID 相同但内容不同）则分配新 ID
     *   2) 遍历源 num，处理 numId 冲突（冲突时分配新 ID）；abstractNumId 按 abs_id_map 重映射
     *   3) 返回 numId 重映射表 {src_numId: dst_numId}，供 copyAllContent 的 applyNumIdMap 使用
     */
    private Map<String, String> mergeNumbering(XWPFDocument src, XWPFDocument dst) {
        Map<String, String> numIdMap = new HashMap<>();
        try {
            XWPFNumbering srcNum = src.getNumbering();
            if (srcNum == null) return numIdMap;
            XWPFNumbering dstNum = dst.getNumbering();
            if (dstNum == null) {
                dstNum = dst.createNumbering();
                if (dstNum == null) return numIdMap;
            }

            // 现有 abstractNum: id -> 序列化内容
            Map<BigInteger, String> dstAbs = new LinkedHashMap<>();
            int maxAbs = 0;
            for (XWPFAbstractNum a : dstNum.getAbstractNums()) {
                CTAbstractNum ct = a.getCTAbstractNum();
                BigInteger id = ct.getAbstractNumId();
                dstAbs.put(id, serializeSafe(ct));
                maxAbs = Math.max(maxAbs, id.intValue());
            }

            Map<BigInteger, BigInteger> absIdMap = new HashMap<>();
            for (XWPFAbstractNum a : srcNum.getAbstractNums()) {
                CTAbstractNum ct = a.getCTAbstractNum();
                BigInteger sid = ct.getAbstractNumId();
                if (!dstAbs.containsKey(sid)) {
                    CTAbstractNum copy = (CTAbstractNum) ct.copy();
                    dstNum.addAbstractNum(new XWPFAbstractNum(copy));
                    dstAbs.put(sid, serializeSafe(copy));
                    absIdMap.put(sid, sid);
                } else if (!serializeSafe(ct).equals(dstAbs.get(sid))) {
                    maxAbs++;
                    BigInteger newId = BigInteger.valueOf(maxAbs);
                    CTAbstractNum copy = (CTAbstractNum) ct.copy();
                    copy.setAbstractNumId(newId);
                    dstNum.addAbstractNum(new XWPFAbstractNum(copy));
                    dstAbs.put(newId, serializeSafe(copy));
                    absIdMap.put(sid, newId);
                } else {
                    absIdMap.put(sid, sid);
                }
            }

            // 现有 num: numId -> abstractNumId
            Map<BigInteger, BigInteger> dstNums = new LinkedHashMap<>();
            int maxNum = 0;
            for (XWPFNum n : dstNum.getNums()) {
                CTNum ct = n.getCTNum();
                BigInteger nid = ct.getNumId();
                BigInteger aid = ct.getAbstractNumId() != null ? ct.getAbstractNumId().getVal() : null;
                dstNums.put(nid, aid);
                maxNum = Math.max(maxNum, nid.intValue());
            }

            for (XWPFNum n : srcNum.getNums()) {
                CTNum ct = n.getCTNum();
                BigInteger snid = ct.getNumId();
                BigInteger said = ct.getAbstractNumId() != null ? ct.getAbstractNumId().getVal() : null;
                BigInteger newAbsId = absIdMap.getOrDefault(said, said);
                if (dstNums.containsKey(snid)) {
                    maxNum++;
                    BigInteger newNid = BigInteger.valueOf(maxNum);
                    CTNum copy = (CTNum) ct.copy();
                    copy.setNumId(newNid);
                    if (copy.getAbstractNumId() != null) copy.getAbstractNumId().setVal(newAbsId);
                    else copy.addNewAbstractNumId().setVal(newAbsId);
                    dstNum.addNum(new XWPFNum(copy));
                    numIdMap.put(snid.toString(), newNid.toString());
                    dstNums.put(newNid, newAbsId);
                } else {
                    CTNum copy = (CTNum) ct.copy();
                    if (copy.getAbstractNumId() != null && newAbsId != null) copy.getAbstractNumId().setVal(newAbsId);
                    else if (newAbsId != null) copy.addNewAbstractNumId().setVal(newAbsId);
                    dstNum.addNum(new XWPFNum(copy));
                    dstNums.put(snid, newAbsId);
                    if (newAbsId != null && !newAbsId.equals(said)) numIdMap.put(snid.toString(), snid.toString());
                }
            }
        } catch (Exception ignore) {
            // 编号合并失败不应中断合并主流程
        }
        return numIdMap;
    }

    private String serializeSafe(CTAbstractNum ct) {
        try {
            return serialize((Element) ct.getDomNode());
        } catch (Exception e) {
            return "";
        }
    }

    private String serialize(Element el) {
        try {
            Transformer t = TransformerFactory.newInstance().newTransformer();
            t.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
            StringWriter sw = new StringWriter();
            t.transform(new DOMSource(el), new StreamResult(sw));
            return sw.toString();
        } catch (Exception e) {
            return "";
        }
    }

    /** 沿祖先链查找 w:tbl，判断段落是否在表格内 */

    private String md5Hex(byte[] data) {
        try {
            MessageDigest md = MessageDigest.getInstance("MD5");
            byte[] d = md.digest(data);
            StringBuilder sb = new StringBuilder();
            for (byte b : d) sb.append(String.format("%02x", b));
            return sb.toString();
        } catch (Exception e) {
            return Integer.toHexString(Arrays.hashCode(data));
        }
    }
}
