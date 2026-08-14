package com.example.project.service.hr.hyp;

import org.w3c.dom.*;
import java.util.List;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Word 合并相关的 OOXML DOM 通用工具方法。
 *
 * <p>从 {@link WordMergeService} 中抽取的零业务耦合的 XML/w3c.dom 操作辅助方法，
 * 便于独立复用与单测。方法均为无状态的静态方法；命名空间常量 {@link #W_NS} 为本类自持，
 * 与 WordMergeService 中同名常量值相同但相互独立（避免跨类依赖）。
 */
public final class OoxmlDomUtils {

    private OoxmlDomUtils() {}

    // ===================== OOXML 命名空间 =====================
    public static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /** 匹配并去除段落开头的旧编号前缀（第X章 /（N）/ N. 等任意格式），返回剩余文本 */
    private static final Pattern OLD_PREFIX = Pattern.compile(
            "^(第[一二三四五六七八九十百零\\d]+章|[（(][一二三四五六七八九十百零\\d]+[)）]|\\d+[\\.、．])\\s*");

    // ===================== 工具：DOM =====================
    public static void walkAndAct(Node node, String ns, String local, Consumer<Element> action) {
        if (node == null) return;
        if (node.getNodeType() == Node.ELEMENT_NODE) {
            Element e = (Element) node;
            if (ns.equals(e.getNamespaceURI()) && local.equals(e.getLocalName())) action.accept(e);
            NodeList children = e.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) walkAndAct(children.item(i), ns, local, action);
        }
    }

    public static Element childEl(Element parent, String ns, String local) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node n = children.item(i);
            if (n.getNodeType() == Node.ELEMENT_NODE) {
                Element e = (Element) n;
                if (ns.equals(e.getNamespaceURI()) && local.equals(e.getLocalName())) return e;
            }
        }
        return null;
    }

    public static Element lastElementChild(Node parent) {
        NodeList kids = parent.getChildNodes();
        for (int i = kids.getLength() - 1; i >= 0; i--) {
            Node n = kids.item(i);
            if (n.getNodeType() == Node.ELEMENT_NODE) return (Element) n;
        }
        return null;
    }

    public static void collectAllElems(Node node, String local, List<Element> out) {
        NodeList kids = node.getChildNodes();
        for (int i = 0; i < kids.getLength(); i++) {
            Node n = kids.item(i);
            if (n.getNodeType() == Node.ELEMENT_NODE) {
                Element e = (Element) n;
                if (W_NS.equals(e.getNamespaceURI()) && local.equals(e.getLocalName())) out.add(e);
                collectAllElems(e, local, out);
            }
        }
    }

    public static String getPStyleVal(Element pEl) {
        Element pPr = childEl(pEl, W_NS, "pPr");
        if (pPr == null) return null;
        Element ps = childEl(pPr, W_NS, "pStyle");
        if (ps == null) return null;
        return ps.getAttributeNS(W_NS, "val");
    }

    /**
     * 获取段落内所有文本内容。
     */
    public static String getFullText(Node node) {
        StringBuilder sb = new StringBuilder();
        walkAndAct(node, W_NS, "t", tNode -> {
            Node child = tNode.getFirstChild();
            while (child != null) {
                if (child.getNodeType() == Node.TEXT_NODE) {
                    sb.append(child.getNodeValue());
                }
                child = child.getNextSibling();
            }
        });
        return sb.toString();
    }

    public static void prependText(Node node, String prefix) {
        Node firstT = firstElementNS(node, W_NS, "t");
        if (firstT != null) {
            Node child = firstT.getFirstChild();
            if (child != null && child.getNodeType() == Node.TEXT_NODE) {
                child.setNodeValue(prefix + child.getNodeValue());
            } else {
                org.w3c.dom.Text txt = firstT.getOwnerDocument().createTextNode(prefix);
                firstT.insertBefore(txt, firstT.getFirstChild());
            }
        } else {
            Element p = (Element) node;
            org.w3c.dom.Document doc = p.getOwnerDocument();
            Element r = doc.createElementNS(W_NS, "w:r");
            Element t = doc.createElementNS(W_NS, "w:t");
            t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
            t.appendChild(doc.createTextNode(prefix));
            r.appendChild(t);
            p.appendChild(r);
        }
    }

    /**
     * 将段落首文本开头的旧编号前缀替换为 newPrefix（先剥离已有编号，再拼新前缀），
     * 避免「第X章 （一）xxx」这类叠加。旧前缀位于第一个 w:t 节点开头（绝大多数情况）。
     */
    public static void setTextPrefix(Element pEl, String newPrefix) {
        Element firstT = (Element) firstElementNS(pEl, W_NS, "t");
        if (firstT != null) {
            Node child = firstT.getFirstChild();
            if (child != null && child.getNodeType() == Node.TEXT_NODE) {
                String old = child.getNodeValue() != null ? child.getNodeValue() : "";
                child.setNodeValue(newPrefix + stripOldPrefix(old));
            } else {
                firstT.appendChild(firstT.getOwnerDocument().createTextNode(newPrefix));
            }
        } else {
            org.w3c.dom.Document doc = pEl.getOwnerDocument();
            Element r = doc.createElementNS(W_NS, "w:r");
            Element t = doc.createElementNS(W_NS, "w:t");
            t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
            t.appendChild(doc.createTextNode(newPrefix));
            r.appendChild(t);
            pEl.appendChild(r);
        }
    }

    public static String stripOldPrefix(String text) {
        if (text == null) return "";
        Matcher m = OLD_PREFIX.matcher(text);
        return m.find() ? text.substring(m.end()) : text;
    }

    public static Node firstElementNS(Node root, String ns, String local) {
        if (root == null) return null;
        if (root.getNodeType() == Node.ELEMENT_NODE) {
            Element e = (Element) root;
            if (ns.equals(e.getNamespaceURI()) && local.equals(e.getLocalName())) return e;
            NodeList children = e.getChildNodes();
            for (int i = 0; i < children.getLength(); i++) {
                Node found = firstElementNS(children.item(i), ns, local);
                if (found != null) return found;
            }
        }
        return null;
    }
}
