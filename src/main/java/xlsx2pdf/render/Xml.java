package xlsx2pdf.render;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Element;
import org.w3c.dom.Node;

/** 名前空間を問わない (ローカル名で照合する) 小さな DOM ヘルパー。 */
public final class Xml {
    private Xml() {}

    /**
     * XML を解析してルート要素を返す。DOCTYPE と外部実体は無効にする (XXE 対策)。
     *
     * @param in XML の入力
     * @return ルート要素
     * @throws Exception 解析に失敗したとき
     */
    public static Element parse(InputStream in) throws Exception {
        DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
        f.setNamespaceAware(true);
        f.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
        f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        f.setExpandEntityReferences(false);
        DocumentBuilder b = f.newDocumentBuilder();
        return b.parse(in).getDocumentElement();
    }

    /** 子要素をすべて返す (e が null なら空)。 */
    public static List<Element> children(Element e) {
        List<Element> out = new ArrayList<>();
        if (e == null) return out;
        for (Node n = e.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n instanceof Element el) out.add(el);
        }
        return out;
    }

    /** ローカル名が一致する子要素をすべて返す (e が null なら空)。 */
    public static List<Element> children(Element e, String local) {
        List<Element> out = new ArrayList<>();
        if (e == null) return out;
        for (Node n = e.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n instanceof Element el && local.equals(el.getLocalName())) out.add(el);
        }
        return out;
    }

    /** ローカル名が一致する最初の子要素を返す (なければ null)。 */
    public static Element child(Element e, String local) {
        if (e == null) return null;
        for (Node n = e.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n instanceof Element el && local.equals(el.getLocalName())) return el;
        }
        return null;
    }

    /** ローカル名の並びをたどって子孫要素を返す (途中でなければ null)。 */
    public static Element path(Element e, String... names) {
        for (String n : names) {
            e = child(e, n);
            if (e == null) return null;
        }
        return e;
    }

    /** 最初の子要素を返す (なければ null)。 */
    public static Element firstChild(Element e) {
        if (e == null) return null;
        for (Node n = e.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n instanceof Element el) return el;
        }
        return null;
    }

    /** 属性値を返す (要素または属性がなければ dflt)。 */
    public static String attr(Element e, String name, String dflt) {
        if (e == null || !e.hasAttribute(name)) return dflt;
        return e.getAttribute(name);
    }

    /** 整数の属性値を返す。小数は四捨五入し、解釈できなければ dflt。 */
    public static long longAttr(Element e, String name, long dflt) {
        if (e == null || !e.hasAttribute(name)) return dflt;
        try {
            return Long.parseLong(e.getAttribute(name).trim());
        } catch (NumberFormatException ex) {
            try {
                return Math.round(Double.parseDouble(e.getAttribute(name).trim()));
            } catch (NumberFormatException ex2) {
                return dflt;
            }
        }
    }

    /** int の属性値を返す ({@link #longAttr} と同じ規則)。 */
    public static int intAttr(Element e, String name, int dflt) {
        return (int) longAttr(e, name, dflt);
    }

    /** 実数の属性値を返す (解釈できなければ dflt)。 */
    public static double doubleAttr(Element e, String name, double dflt) {
        if (e == null || !e.hasAttribute(name)) return dflt;
        try {
            return Double.parseDouble(e.getAttribute(name).trim());
        } catch (NumberFormatException ex) {
            return dflt;
        }
    }

    /** 真偽値の属性値を返す ("1"・"true"・"t" を真とする)。 */
    public static boolean boolAttr(Element e, String name, boolean dflt) {
        if (e == null || !e.hasAttribute(name)) return dflt;
        String v = e.getAttribute(name);
        return v.equals("1") || v.equalsIgnoreCase("true") || v.equalsIgnoreCase("t");
    }

    /** 要素の文字列内容を返す (null なら空文字列)。 */
    public static String text(Element e) {
        return e == null ? "" : e.getTextContent();
    }

    /**
     * mc:AlternateContent を解決した子要素の並びを返す。要件 (Requires) をすべて理解できる Choice を優先し、
     * なければ Fallback を使う。AlternateContent 以外の子要素はそのまま含める。
     *
     * @param e         親要素
     * @param supported 理解できる名前空間接頭辞
     */
    public static List<Element> resolveAlternate(Element e, java.util.Set<String> supported) {
        List<Element> out = new ArrayList<>();
        for (Element c : children(e)) {
            if ("AlternateContent".equals(c.getLocalName())) {
                Element chosen = null;
                for (Element ch : children(c)) {
                    if ("Choice".equals(ch.getLocalName())) {
                        String req = ch.getAttribute("Requires");
                        boolean ok = true;
                        for (String r : req.split("\\s+")) if (!r.isEmpty() && !supported.contains(r)) ok = false;
                        if (ok) { chosen = ch; break; }
                    } else if ("Fallback".equals(ch.getLocalName()) && chosen == null) {
                        chosen = ch;
                    }
                }
                if (chosen != null) out.addAll(children(chosen));
            } else {
                out.add(c);
            }
        }
        return out;
    }
}
