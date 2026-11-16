package xlsx2pdf.render;

import java.awt.Color;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Arc2D;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.geom.RoundRectangle2D;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 * SVG の画像 (asvg:svgBlip) をベクターのまま描く。Excel と同じく、viewBox (なければ 0 0 width height) を画像の枠に
 * 縦横別々に合わせる (preserveAspectRatio は使わない)。
 * <p>
 * 対応: svg・g・defs・symbol・use・path (全命令)・rect (角の丸め)・circle・ellipse・line・polyline・polygon・text・tspan、
 * transform、表現属性と style 属性と &lt;style&gt; の CSS (要素名・.class・#id とその組み合わせ。子孫の結合子は使わない)、
 * fill・stroke (色、none、currentColor、url(#線形・放射状のグラデーション))、各種の不透明度 (グループの opacity は子に掛ける)、
 * fill-rule、線の端・角・破線、clip-path、display・visibility。対応しない要素 (image・filter・mask・marker など) は描かない。
 */
final class SvgPainter {
    /** 参照 (use、グラデーションの href) をたどる深さの上限 (循環の防止)。 */
    private static final int MAX_DEPTH = 32;

    private final RenderContext ctx;
    private final Map<String, Element> ids = new HashMap<>();
    private final List<Rule> rules = new ArrayList<>();
    private double vbW = 100, vbH = 100;
    private int depth;

    /** CSS の規則 1 つ (単純な選択子 1 つと宣言)。 */
    private record Rule(String tag, String id, List<String> classes, int specificity, int order, Map<String, String> decl) {}

    private SvgPainter(RenderContext ctx) { this.ctx = ctx; }

    /**
     * SVG を矩形 r に描く。
     *
     * @return 描いたか (SVG として読めなければ false。呼び出し側は代わりの画像を描く)
     */
    static boolean paint(RenderContext ctx, PdfCanvas cv, byte[] svg, Rectangle2D r) {
        Element root;
        try {
            root = Xml.parse(new ByteArrayInputStream(svg));
        } catch (Exception e) {
            return false;
        }
        if (root == null || !"svg".equals(root.getLocalName())) return false;
        SvgPainter p = new SvgPainter(ctx);
        double[] vb = p.viewBox(root);
        if (vb == null) return false;
        p.vbW = vb[2];
        p.vbH = vb[3];
        p.index(root);
        try {
            cv.save();
            cv.clipRect(r.getX(), r.getY(), r.getWidth(), r.getHeight());
            AffineTransform at = AffineTransform.getTranslateInstance(r.getX(), r.getY());
            at.scale(r.getWidth() / vb[2], r.getHeight() / vb[3]);
            at.translate(-vb[0], -vb[1]);
            cv.transform(at);
            p.children(cv, root, Style.ROOT);
            cv.restore();
            return true;
        } catch (IOException | RuntimeException e) {
            Warnings.add("SVG の画像を描けませんでした", e);
            return false;
        }
    }

    /** ルートの viewBox (なければ 0 0 width height)。大きさが分からなければ null。 */
    private double[] viewBox(Element svg) {
        double[] vb = numbers(svg.getAttribute("viewBox"));
        if (vb.length == 4 && vb[2] > 0 && vb[3] > 0) return vb;
        double w = length(svg.getAttribute("width"), 0, 0), h = length(svg.getAttribute("height"), 0, 0);
        if (w > 0 && h > 0) return new double[] {0, 0, w, h};
        return null;
    }

    // ------------------------------------------------------------ 索引と CSS

    private void index(Element e) {
        if (!e.getAttribute("id").isEmpty()) ids.putIfAbsent(e.getAttribute("id"), e);
        if ("style".equals(e.getLocalName())) parseCss(e.getTextContent());
        for (Element c : Xml.children(e)) index(c);
    }

    private void parseCss(String css) {
        css = css.replaceAll("(?s)/\\*.*?\\*/", "");
        int i = 0;
        while (i < css.length()) {
            int open = css.indexOf('{', i), close = open < 0 ? -1 : css.indexOf('}', open);
            if (open < 0 || close < 0) break;
            String sel = css.substring(i, open).trim();
            Map<String, String> decl = declarations(css.substring(open + 1, close));
            i = close + 1;
            if (sel.startsWith("@")) continue;
            for (String one : sel.split(",")) {
                String s = one.trim();
                if (s.isEmpty() || s.matches(".*[\\s>+~\\[:].*")) continue;   // 結合子・属性・疑似クラスは使わない
                String tag = null, id = null;
                List<String> cls = new ArrayList<>();
                java.util.regex.Matcher m = java.util.regex.Pattern.compile("([.#]?)([^.#]+)").matcher(s);
                while (m.find()) {
                    switch (m.group(1)) {
                        case "." -> cls.add(m.group(2));
                        case "#" -> id = m.group(2);
                        default -> tag = m.group(2).equals("*") ? null : m.group(2);
                    }
                }
                int spec = (id != null ? 100 : 0) + cls.size() * 10 + (tag != null ? 1 : 0);
                rules.add(new Rule(tag, id, cls, spec, rules.size(), decl));
            }
        }
        rules.sort((a, b) -> a.specificity != b.specificity ? Integer.compare(a.specificity, b.specificity) : Integer.compare(a.order, b.order));
    }

    private static Map<String, String> declarations(String s) {
        Map<String, String> m = new HashMap<>();
        for (String d : s.split(";")) {
            int c = d.indexOf(':');
            if (c <= 0) continue;
            String v = d.substring(c + 1).trim().replace("!important", "").trim();
            m.put(d.substring(0, c).trim().toLowerCase(Locale.ROOT), v);
        }
        return m;
    }

    /** 要素の宣言 (表現属性 → CSS の規則 → style 属性の順に上書き)。 */
    private Map<String, String> declared(Element e) {
        Map<String, String> m = new HashMap<>();
        var attrs = e.getAttributes();
        for (int i = 0; i < attrs.getLength(); i++) {
            Node a = attrs.item(i);
            String n = a.getLocalName() != null ? a.getLocalName() : a.getNodeName();
            if (Style.PROPS.contains(n)) m.put(n, a.getNodeValue().trim());
        }
        if (!rules.isEmpty()) {
            List<String> cls = List.of(e.getAttribute("class").trim().split("\\s+"));
            for (Rule r : rules) {
                if (r.tag != null && !r.tag.equals(e.getLocalName())) continue;
                if (r.id != null && !r.id.equals(e.getAttribute("id"))) continue;
                if (!cls.containsAll(r.classes)) continue;
                m.putAll(r.decl);
            }
        }
        if (e.hasAttribute("style")) m.putAll(declarations(e.getAttribute("style")));
        return m;
    }

    // ------------------------------------------------------------ 継承するスタイル

    /** 描画の状態 (継承する値と、要素ごとに掛けた不透明度)。 */
    private static final class Style {
        static final java.util.Set<String> PROPS = java.util.Set.of("fill", "stroke", "stroke-width", "fill-opacity", "stroke-opacity",
                "opacity", "fill-rule", "clip-rule", "stroke-linecap", "stroke-linejoin", "stroke-miterlimit", "stroke-dasharray",
                "stroke-dashoffset", "font-family", "font-size", "font-weight", "font-style", "text-anchor", "visibility", "display",
                "color", "clip-path", "stop-color", "stop-opacity");
        static final Style ROOT = new Style();

        String fill = "black", stroke = "none", fillRule = "nonzero", clipRule = "nonzero", cap = "butt", join = "miter",
                dash = "none", fontFamily = null, fontWeight = "normal", fontStyle = "normal", anchor = "start", visibility = "visible";
        Color color = Color.BLACK;
        double strokeWidth = 1, fillOpacity = 1, strokeOpacity = 1, groupOpacity = 1, fontSize = 16;

        Style copy() {
            Style s = new Style();
            s.fill = fill; s.stroke = stroke; s.fillRule = fillRule; s.clipRule = clipRule; s.cap = cap; s.join = join; s.dash = dash;
            s.fontFamily = fontFamily; s.fontWeight = fontWeight; s.fontStyle = fontStyle; s.anchor = anchor; s.visibility = visibility;
            s.color = color; s.strokeWidth = strokeWidth; s.fillOpacity = fillOpacity; s.strokeOpacity = strokeOpacity;
            s.groupOpacity = groupOpacity; s.fontSize = fontSize;
            return s;
        }
    }

    /** 親の状態に要素の宣言を当てた状態。 */
    private Style style(Element e, Style parent, Map<String, String> d) {
        Style s = parent.copy();
        String v;
        if ((v = d.get("color")) != null && !v.equals("inherit")) { Color c = color(v, s.color); if (c != null) s.color = c; }
        if ((v = d.get("fill")) != null && !v.equals("inherit")) s.fill = v;
        if ((v = d.get("stroke")) != null && !v.equals("inherit")) s.stroke = v;
        if ((v = d.get("stroke-width")) != null && !v.equals("inherit")) s.strokeWidth = length(v, s.strokeWidth, diag());
        if ((v = d.get("fill-opacity")) != null) s.fillOpacity = opacity(v, s.fillOpacity);
        if ((v = d.get("stroke-opacity")) != null) s.strokeOpacity = opacity(v, s.strokeOpacity);
        if ((v = d.get("opacity")) != null) s.groupOpacity *= opacity(v, 1);
        if ((v = d.get("fill-rule")) != null && !v.equals("inherit")) s.fillRule = v;
        if ((v = d.get("clip-rule")) != null && !v.equals("inherit")) s.clipRule = v;
        if ((v = d.get("stroke-linecap")) != null && !v.equals("inherit")) s.cap = v;
        if ((v = d.get("stroke-linejoin")) != null && !v.equals("inherit")) s.join = v;
        if ((v = d.get("stroke-dasharray")) != null && !v.equals("inherit")) s.dash = v;
        if ((v = d.get("font-family")) != null && !v.equals("inherit")) s.fontFamily = v;
        if ((v = d.get("font-size")) != null && !v.equals("inherit")) s.fontSize = length(v, s.fontSize, s.fontSize);
        if ((v = d.get("font-weight")) != null && !v.equals("inherit")) s.fontWeight = v;
        if ((v = d.get("font-style")) != null && !v.equals("inherit")) s.fontStyle = v;
        if ((v = d.get("text-anchor")) != null && !v.equals("inherit")) s.anchor = v;
        if ((v = d.get("visibility")) != null && !v.equals("inherit")) s.visibility = v;
        return s;
    }

    // ------------------------------------------------------------ 要素

    private void children(PdfCanvas cv, Element e, Style st) throws IOException {
        for (Element c : Xml.children(e)) element(cv, c, st);
    }

    private void element(PdfCanvas cv, Element e, Style parent) throws IOException {
        String tag = e.getLocalName();
        if (tag == null) return;
        switch (tag) {
            case "defs", "style", "title", "desc", "metadata", "clipPath", "mask", "linearGradient", "radialGradient",
                 "pattern", "marker", "filter", "symbol", "script" -> { return; }
            default -> { }
        }
        Map<String, String> d = declared(e);
        if ("none".equals(d.get("display"))) return;
        Style st = style(e, parent, d);
        AffineTransform tf = transform(e.getAttribute("transform"));
        Shape clip = clip(d.get("clip-path"), e);
        cv.save();
        if (tf != null) cv.transform(tf);
        if (clip != null) cv.clip(clip);
        try {
            switch (tag) {
                case "g", "a", "switch" -> children(cv, e, st);
                case "svg" -> nested(cv, e, st);
                case "use" -> use(cv, e, st);
                case "text" -> text(cv, e, st);
                default -> {
                    Shape s = shape(e);
                    if (s != null && "visible".equals(st.visibility)) paintShape(cv, s, st, e);
                }
            }
        } finally {
            cv.restore();
        }
    }

    /** 入れ子の svg: x・y へ移し、viewBox があれば width・height に合わせる。 */
    private void nested(PdfCanvas cv, Element e, Style st) throws IOException {
        double x = length(e.getAttribute("x"), 0, vbW), y = length(e.getAttribute("y"), 0, vbH);
        double w = length(e.getAttribute("width"), vbW, vbW), h = length(e.getAttribute("height"), vbH, vbH);
        AffineTransform at = AffineTransform.getTranslateInstance(x, y);
        double[] vb = numbers(e.getAttribute("viewBox"));
        if (vb.length == 4 && vb[2] > 0 && vb[3] > 0) {
            cv.clipRect(x, y, w, h);
            at.scale(w / vb[2], h / vb[3]);
            at.translate(-vb[0], -vb[1]);
        }
        cv.transform(at);
        children(cv, e, st);
    }

    /** use: 参照先を x・y へ移して描く (symbol は中身を描く)。 */
    private void use(PdfCanvas cv, Element e, Style st) throws IOException {
        Element ref = ref(href(e));
        if (ref == null || depth >= MAX_DEPTH) return;
        double x = length(e.getAttribute("x"), 0, vbW), y = length(e.getAttribute("y"), 0, vbH);
        if (x != 0 || y != 0) cv.translate(x, y);
        depth++;
        try {
            if ("symbol".equals(ref.getLocalName())) {
                Style s2 = style(ref, st, declared(ref));
                double[] vb = numbers(ref.getAttribute("viewBox"));
                if (vb.length == 4 && vb[2] > 0 && vb[3] > 0 && e.hasAttribute("width") && e.hasAttribute("height")) {
                    double w = length(e.getAttribute("width"), vb[2], vbW), h = length(e.getAttribute("height"), vb[3], vbH);
                    AffineTransform at = AffineTransform.getScaleInstance(w / vb[2], h / vb[3]);
                    at.translate(-vb[0], -vb[1]);
                    cv.transform(at);
                }
                children(cv, ref, s2);
            } else {
                element(cv, ref, st);
            }
        } finally {
            depth--;
        }
    }

    private static String href(Element e) {
        String h = e.getAttributeNS("http://www.w3.org/1999/xlink", "href");
        if (h.isEmpty()) h = e.getAttribute("href");
        if (h.isEmpty()) h = e.getAttribute("xlink:href");
        return h;
    }

    private Element ref(String h) {
        return h != null && h.startsWith("#") ? ids.get(h.substring(1)) : null;
    }

    /** 図形の要素の形 (なければ null)。 */
    private Shape shape(Element e) {
        switch (e.getLocalName()) {
            case "path" -> { return SvgPath.parse(e.getAttribute("d")); }
            case "rect" -> {
                double x = length(e.getAttribute("x"), 0, vbW), y = length(e.getAttribute("y"), 0, vbH);
                double w = length(e.getAttribute("width"), 0, vbW), h = length(e.getAttribute("height"), 0, vbH);
                if (w <= 0 || h <= 0) return null;
                double rx = length(e.getAttribute("rx"), -1, vbW), ry = length(e.getAttribute("ry"), -1, vbH);
                if (rx < 0) rx = ry;
                if (ry < 0) ry = rx;
                if (rx <= 0) return new Rectangle2D.Double(x, y, w, h);
                return new RoundRectangle2D.Double(x, y, w, h, 2 * Math.min(rx, w / 2), 2 * Math.min(ry, h / 2));
            }
            case "circle" -> {
                double cx = length(e.getAttribute("cx"), 0, vbW), cy = length(e.getAttribute("cy"), 0, vbH), r = length(e.getAttribute("r"), 0, diag());
                return r > 0 ? new Ellipse2D.Double(cx - r, cy - r, 2 * r, 2 * r) : null;
            }
            case "ellipse" -> {
                double cx = length(e.getAttribute("cx"), 0, vbW), cy = length(e.getAttribute("cy"), 0, vbH);
                double rx = length(e.getAttribute("rx"), 0, vbW), ry = length(e.getAttribute("ry"), 0, vbH);
                return rx > 0 && ry > 0 ? new Ellipse2D.Double(cx - rx, cy - ry, 2 * rx, 2 * ry) : null;
            }
            case "line" -> {
                Path2D p = new Path2D.Double();
                p.moveTo(length(e.getAttribute("x1"), 0, vbW), length(e.getAttribute("y1"), 0, vbH));
                p.lineTo(length(e.getAttribute("x2"), 0, vbW), length(e.getAttribute("y2"), 0, vbH));
                return p;
            }
            case "polyline", "polygon" -> {
                double[] v = numbers(e.getAttribute("points"));
                if (v.length < 4) return null;
                Path2D p = new Path2D.Double();
                p.moveTo(v[0], v[1]);
                for (int i = 2; i + 1 < v.length; i += 2) p.lineTo(v[i], v[i + 1]);
                if ("polygon".equals(e.getLocalName())) p.closePath();
                return p;
            }
            default -> { return null; }
        }
    }

    /** 塗りと線を描く。 */
    private void paintShape(PdfCanvas cv, Shape s, Style st, Element e) throws IOException {
        boolean open = "line".equals(e.getLocalName()) || "polyline".equals(e.getLocalName());
        if (!open) {
            Path2D p = new Path2D.Double(s);
            p.setWindingRule("evenodd".equals(st.fillRule) ? Path2D.WIND_EVEN_ODD : Path2D.WIND_NON_ZERO);
            paint(cv, p, st.fill, st.fillOpacity * st.groupOpacity, st, true);
        } else if (!"none".equals(st.fill) && "polyline".equals(e.getLocalName())) {
            paint(cv, s, st.fill, st.fillOpacity * st.groupOpacity, st, true);
        }
        if (!"none".equals(st.stroke) && st.strokeWidth > 0) paint(cv, s, st.stroke, st.strokeOpacity * st.groupOpacity, st, false);
    }

    /** 色・グラデーションで塗る (fill) か線を描く。 */
    private void paint(PdfCanvas cv, Shape s, String spec, double alpha, Style st, boolean fill) throws IOException {
        if (spec == null || spec.equals("none") || alpha <= 0) return;
        if (spec.startsWith("url(")) {
            int end = spec.indexOf(')');
            String id = spec.substring(4, end < 0 ? spec.length() : end).trim().replace("\"", "").replace("'", "");
            Element g = ref(id);
            if (g != null && fill && gradient(cv, s, g, alpha)) return;
            String fb = end >= 0 ? spec.substring(end + 1).trim() : "";
            if (g != null && !fill) {   // 線のグラデーションは最初の色で描く
                List<Object[]> stops = stops(g, 0);
                if (!stops.isEmpty()) { strokeOrFill(cv, s, (Color) stops.get(0)[1], alpha, st, false); return; }
            }
            if (fb.isEmpty() || fb.startsWith("url(")) return;
            spec = fb;
        }
        Color c = color(spec, st.color);
        if (c != null) strokeOrFill(cv, s, c, alpha, st, fill);
    }

    private void strokeOrFill(PdfCanvas cv, Shape s, Color c, double alpha, Style st, boolean fill) throws IOException {
        int a = (int) Math.round(Math.max(0, Math.min(1, alpha * c.getAlpha() / 255.0)) * 255);
        Color col = new Color(c.getRed(), c.getGreen(), c.getBlue(), a);
        if (fill) {
            cv.fill(s, col);
            return;
        }
        PdfCanvas.Stroke k = new PdfCanvas.Stroke((float) st.strokeWidth);
        k.cap = switch (st.cap) { case "round" -> 1; case "square" -> 2; default -> 0; };
        k.join = switch (st.join) { case "round" -> 1; case "bevel" -> 2; default -> 0; };
        if (!"none".equals(st.dash)) {
            double[] v = numbers(st.dash.replace("px", ""));
            if (v.length % 2 == 1) { double[] w = new double[v.length * 2]; System.arraycopy(v, 0, w, 0, v.length); System.arraycopy(v, 0, w, v.length, v.length); v = w; }
            boolean any = false;
            float[] f = new float[v.length];
            for (int i = 0; i < v.length; i++) { f[i] = (float) Math.max(0, v[i]); any |= f[i] > 0; }
            if (any) k.dash = f;
        }
        cv.stroke(s, k, col);
    }

    // ------------------------------------------------------------ グラデーション

    /** グラデーションで図形を塗る。読めなければ false。 */
    private boolean gradient(PdfCanvas cv, Shape s, Element g, double alpha) throws IOException {
        List<Object[]> stops = stops(g, 0);
        if (stops.isEmpty()) return false;
        double[] pos = new double[stops.size()];
        Color[] cols = new Color[stops.size()];
        double prev = 0;
        for (int i = 0; i < pos.length; i++) {
            prev = Math.max(prev, (double) stops.get(i)[0]);
            pos[i] = prev;
            Color c = (Color) stops.get(i)[1];
            cols[i] = c;
        }
        if (pos.length == 1) {
            Color c = cols[0];
            cv.fill(s, new Color(c.getRed(), c.getGreen(), c.getBlue(), (int) Math.round(alpha * c.getAlpha())));
            return true;
        }
        boolean user = "userSpaceOnUse".equals(gAttr(g, "gradientUnits", 0));
        Rectangle2D bb = s.getBounds2D();
        if (!user && (bb.getWidth() <= 0 || bb.getHeight() <= 0)) return false;
        AffineTransform at = new AffineTransform();
        if (!user) { at.translate(bb.getX(), bb.getY()); at.scale(bb.getWidth(), bb.getHeight()); }
        AffineTransform gt = transform(gAttr(g, "gradientTransform", 0));
        if (gt != null) at.concatenate(gt);
        double rw = user ? vbW : 1, rh = user ? vbH : 1;
        cv.save();
        cv.clip(s);
        cv.transform(at);
        if ("radialGradient".equals(g.getLocalName())) {
            double cx = length(gAttr(g, "cx", 0), user ? vbW / 2 : 0.5, rw), cy = length(gAttr(g, "cy", 0), user ? vbH / 2 : 0.5, rh);
            double r = length(gAttr(g, "r", 0), user ? diag() / 2 : 0.5, user ? diag() : 1);
            String fxs = gAttr(g, "fx", 0), fys = gAttr(g, "fy", 0);
            double fx = fxs == null ? cx : length(fxs, cx, rw), fy = fys == null ? cy : length(fys, cy, rh);
            cv.fillRadialStops(fx, fy, cx, cy, r, pos, cols);
        } else {
            double x1 = length(gAttr(g, "x1", 0), 0, rw), y1 = length(gAttr(g, "y1", 0), 0, rh);
            double x2 = length(gAttr(g, "x2", 0), user ? vbW : 1, rw), y2 = length(gAttr(g, "y2", 0), 0, rh);
            cv.fillAxialStops(x1, y1, x2, y2, pos, cols);
        }
        cv.restore();
        return true;
    }

    /** グラデーションの属性 (なければ href の先から)。なければ null。 */
    private String gAttr(Element g, String name, int d) {
        if (g == null || d > MAX_DEPTH) return null;
        if (g.hasAttribute(name)) return g.getAttribute(name);
        return gAttr(ref(href(g)), name, d + 1);
    }

    /** 色の位置 {offset, Color} (自分になければ href の先の stop)。 */
    private List<Object[]> stops(Element g, int d) {
        List<Object[]> out = new ArrayList<>();
        if (g == null || d > MAX_DEPTH) return out;
        for (Element s : Xml.children(g)) {
            if (!"stop".equals(s.getLocalName())) continue;
            Map<String, String> m = declared(s);
            String o = s.getAttribute("offset").trim();
            double off = o.endsWith("%") ? num(o.substring(0, o.length() - 1), 0) / 100 : num(o, 0);
            Color c = color(m.getOrDefault("stop-color", "black"), Color.BLACK);
            if (c == null) c = Color.BLACK;
            double a = opacity(m.getOrDefault("stop-opacity", "1"), 1) * c.getAlpha() / 255.0;
            out.add(new Object[] {Math.max(0, Math.min(1, off)), new Color(c.getRed(), c.getGreen(), c.getBlue(), (int) Math.round(a * 255))});
        }
        return out.isEmpty() ? stops(ref(href(g)), d + 1) : out;
    }

    // ------------------------------------------------------------ クリップ

    /** clip-path="url(#id)" の形 (clipPath の子の形を合わせたもの。参照する要素の座標)。なければ null。 */
    private Shape clip(String v, Element owner) {
        if (v == null || !v.startsWith("url(")) return null;
        int end = v.indexOf(')');
        Element cp = ref(v.substring(4, end < 0 ? v.length() : end).trim().replace("\"", "").replace("'", ""));
        if (cp == null || !"clipPath".equals(cp.getLocalName())) return null;
        Path2D all = new Path2D.Double();
        boolean bbox = "objectBoundingBox".equals(cp.getAttribute("clipPathUnits"));
        for (Element c : Xml.children(cp)) {
            Element el = "use".equals(c.getLocalName()) ? ref(href(c)) : c;
            if (el == null) continue;
            Shape s = shape(el);
            if (s == null) continue;
            AffineTransform t = transform(c.getAttribute("transform"));
            if ("use".equals(c.getLocalName())) {
                AffineTransform u = AffineTransform.getTranslateInstance(length(c.getAttribute("x"), 0, vbW), length(c.getAttribute("y"), 0, vbH));
                if (t != null) t.concatenate(u); else t = u;
            }
            all.append(t == null ? s : t.createTransformedShape(s), false);
        }
        AffineTransform ct = transform(cp.getAttribute("transform"));
        Shape out = ct == null ? all : ct.createTransformedShape(all);
        if (bbox) {
            Shape os = shape(owner);
            if (os == null) return null;
            Rectangle2D bb = os.getBounds2D();
            AffineTransform b = AffineTransform.getTranslateInstance(bb.getX(), bb.getY());
            b.scale(bb.getWidth(), bb.getHeight());
            out = b.createTransformedShape(out);
        }
        return out;
    }

    // ------------------------------------------------------------ 文字

    /** text と tspan: 文字を左から続けて置く (tspan の x・y・dx・dy、text-anchor は行の全体で)。 */
    private void text(PdfCanvas cv, Element e, Style st) throws IOException {
        List<Object[]> runs = new ArrayList<>();   // {String, Style, x or NaN, y or NaN, dx, dy}
        collectText(e, st, runs, true);
        double x = length(first(e.getAttribute("x")), 0, vbW), y = length(first(e.getAttribute("y")), 0, vbH);
        // 塊ごと (x の指定で区切る) に幅を測り、text-anchor で寄せる
        int i = 0;
        while (i < runs.size()) {
            int j = i;
            double w = 0;
            do {
                Object[] r = runs.get(j);
                w += width((String) r[0], (Style) r[1]) + (j > i ? (double) r[4] : 0);
                j++;
            } while (j < runs.size() && Double.isNaN((double) runs.get(j)[2]));
            Object[] head = runs.get(i);
            if (!Double.isNaN((double) head[2])) x = (double) head[2];
            if (!Double.isNaN((double) head[3])) y = (double) head[3];
            x += (double) head[4];
            y += (double) head[5];
            String anchor = ((Style) head[1]).anchor;
            if ("middle".equals(anchor)) x -= w / 2;
            else if ("end".equals(anchor)) x -= w;
            for (int k = i; k < j; k++) {
                Object[] r = runs.get(k);
                Style s = (Style) r[1];
                if (k > i) { x += (double) r[4]; y += (double) r[5]; }
                String t = (String) r[0];
                if (!"visible".equals(s.visibility) || "none".equals(s.fill) || s.fill.startsWith("url(")) { x += width(t, s); continue; }
                Color c = color(s.fill, s.color);
                if (c != null) {
                    double a = s.fillOpacity * s.groupOpacity * c.getAlpha() / 255.0;
                    cv.text(t, x, y, font(s), (float) s.fontSize, new Color(c.getRed(), c.getGreen(), c.getBlue(), (int) Math.round(a * 255)), 1, 0);
                }
                x += width(t, s);
            }
            i = j;
        }
    }

    private void collectText(Element e, Style st, List<Object[]> out, boolean root) {
        boolean firstRun = true;
        for (Node n = e.getFirstChild(); n != null; n = n.getNextSibling()) {
            if (n.getNodeType() == Node.TEXT_NODE || n.getNodeType() == Node.CDATA_SECTION_NODE) {
                String t = n.getNodeValue().replaceAll("[\\r\\n\\t]", " ").replaceAll(" +", " ");
                if (out.isEmpty()) t = t.replaceAll("^ +", "");
                if (n.getNextSibling() == null && root) t = t.replaceAll(" +$", "");
                if (t.isEmpty()) continue;
                boolean head = firstRun && !root;
                out.add(new Object[] {t, st, head ? xy(e, "x", vbW) : Double.NaN, head ? xy(e, "y", vbH) : Double.NaN,
                        head ? length(first(e.getAttribute("dx")), 0, vbW) : 0.0, head ? length(first(e.getAttribute("dy")), 0, vbH) : 0.0});
                firstRun = false;
            } else if (n instanceof Element c && "tspan".equals(c.getLocalName())) {
                Map<String, String> d = declared(c);
                if ("none".equals(d.get("display"))) continue;
                collectText(c, style(c, st, d), out, false);
            }
        }
    }

    private double xy(Element e, String a, double ref) {
        return e.hasAttribute(a) ? length(first(e.getAttribute(a)), 0, ref) : Double.NaN;
    }

    private static String first(String v) {
        String[] p = v.trim().split("[\\s,]+");
        return p.length == 0 ? "" : p[0];
    }

    private FontManager.Spec font(Style s) {
        String fam = s.fontFamily == null ? "" : s.fontFamily.split(",")[0].trim().replace("'", "").replace("\"", "");
        if (fam.isEmpty() || fam.equals("sans-serif")) fam = "Arial";
        else if (fam.equals("serif")) fam = "Times New Roman";
        else if (fam.equals("monospace")) fam = "Courier New";
        boolean bold = s.fontWeight.equals("bold") || s.fontWeight.equals("bolder") || num(s.fontWeight, 400) >= 600;
        boolean italic = s.fontStyle.equals("italic") || s.fontStyle.equals("oblique");
        return ctx.fonts.spec(fam, bold, italic);
    }

    private double width(String t, Style s) {
        return font(s).width(t, (float) s.fontSize, ctx.fonts);
    }

    // ------------------------------------------------------------ 値

    private double diag() { return Math.sqrt((vbW * vbW + vbH * vbH) / 2); }

    /** 長さ (px・pt・mm・cm・in・em・%)。%の基準は ref。空や読めなければ def。 */
    private static double length(String v, double def, double ref) {
        if (v == null) return def;
        v = v.trim();
        if (v.isEmpty()) return def;
        double k = 1;
        String n = v;
        if (v.endsWith("%")) { n = v.substring(0, v.length() - 1); k = ref / 100; }
        else if (v.endsWith("px")) n = v.substring(0, v.length() - 2);
        else if (v.endsWith("pt")) { n = v.substring(0, v.length() - 2); k = 4 / 3.0; }
        else if (v.endsWith("mm")) { n = v.substring(0, v.length() - 2); k = 96 / 25.4; }
        else if (v.endsWith("cm")) { n = v.substring(0, v.length() - 2); k = 96 / 2.54; }
        else if (v.endsWith("in")) { n = v.substring(0, v.length() - 2); k = 96; }
        else if (v.endsWith("em")) { n = v.substring(0, v.length() - 2); k = 16; }
        double d = num(n, Double.NaN);
        return Double.isNaN(d) ? def : d * k;
    }

    private static double num(String s, double def) {
        try {
            return Double.parseDouble(s.trim());
        } catch (NumberFormatException | NullPointerException e) {
            return def;
        }
    }

    private static double opacity(String v, double def) {
        v = v.trim();
        double d = v.endsWith("%") ? num(v.substring(0, v.length() - 1), Double.NaN) / 100 : num(v, Double.NaN);
        return Double.isNaN(d) ? def : Math.max(0, Math.min(1, d));
    }

    /** 数の並び (区切りは空白・カンマ。"1-2" や "1.5.5" のように詰めて書いたものも分ける)。 */
    static double[] numbers(String s) {
        if (s == null) return new double[0];
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("[-+]?(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][-+]?\\d+)?").matcher(s);
        List<Double> out = new ArrayList<>();
        while (m.find()) out.add(Double.parseDouble(m.group()));
        double[] a = new double[out.size()];
        for (int i = 0; i < a.length; i++) a[i] = out.get(i);
        return a;
    }

    /** CSS の色 (#rgb・#rrggbb・#rrggbbaa・rgb()・rgba()・色名・currentColor・transparent)。none や読めなければ null。 */
    static Color color(String v, Color current) {
        if (v == null) return null;
        v = v.trim();
        String lv = v.toLowerCase(Locale.ROOT);
        if (lv.isEmpty() || lv.equals("none")) return null;
        if (lv.equals("currentcolor")) return current;
        if (lv.equals("transparent")) return new Color(0, 0, 0, 0);
        try {
            if (lv.startsWith("#")) {
                String h = lv.substring(1);
                if (h.length() == 3 || h.length() == 4) {
                    StringBuilder b = new StringBuilder();
                    for (char ch : h.toCharArray()) b.append(ch).append(ch);
                    h = b.toString();
                }
                if (h.length() == 6) return new Color(Integer.parseInt(h, 16));
                if (h.length() == 8) return new Color(Integer.parseInt(h.substring(0, 2), 16), Integer.parseInt(h.substring(2, 4), 16),
                        Integer.parseInt(h.substring(4, 6), 16), Integer.parseInt(h.substring(6, 8), 16));
                return null;
            }
            if (lv.startsWith("rgb")) {
                String in = lv.substring(lv.indexOf('(') + 1, lv.lastIndexOf(')'));
                String[] p = in.split("[\\s,/]+");
                int[] c = new int[3];
                for (int i = 0; i < 3; i++) {
                    String q = p[i].trim();
                    c[i] = (int) Math.round(q.endsWith("%") ? num(q.substring(0, q.length() - 1), 0) * 2.55 : num(q, 0));
                    c[i] = Math.max(0, Math.min(255, c[i]));
                }
                int a = p.length > 3 ? (int) Math.round(opacity(p[3], 1) * 255) : 255;
                return new Color(c[0], c[1], c[2], a);
            }
        } catch (RuntimeException e) {
            return null;
        }
        Color n = NAMED.get(lv);
        return n;
    }

    /** CSS の色名 (POI のプリセット色の表を小文字で引く。grey も gray と同じ)。 */
    private static final Map<String, Color> NAMED = new HashMap<>();
    static {
        for (org.apache.poi.sl.usermodel.PresetColor p : org.apache.poi.sl.usermodel.PresetColor.values()) {
            if (p.ooxmlId == null || p.color == null) continue;
            String k = p.ooxmlId.toLowerCase(Locale.ROOT);
            NAMED.putIfAbsent(k, p.color);
            if (k.contains("gray")) NAMED.putIfAbsent(k.replace("gray", "grey"), p.color);
        }
    }

    /** transform 属性 (matrix・translate・scale・rotate・skewX・skewY の並び)。なければ null。 */
    static AffineTransform transform(String v) {
        if (v == null || v.isBlank()) return null;
        AffineTransform at = new AffineTransform();
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("(matrix|translate|scale|rotate|skewX|skewY)\\s*\\(([^)]*)\\)").matcher(v);
        boolean any = false;
        while (m.find()) {
            double[] a = numbers(m.group(2));
            any = true;
            switch (m.group(1)) {
                case "matrix" -> { if (a.length == 6) at.concatenate(new AffineTransform(a[0], a[1], a[2], a[3], a[4], a[5])); }
                case "translate" -> { if (a.length >= 1) at.translate(a[0], a.length > 1 ? a[1] : 0); }
                case "scale" -> { if (a.length >= 1) at.scale(a[0], a.length > 1 ? a[1] : a[0]); }
                case "rotate" -> {
                    if (a.length >= 3) at.rotate(Math.toRadians(a[0]), a[1], a[2]);
                    else if (a.length >= 1) at.rotate(Math.toRadians(a[0]));
                }
                case "skewX" -> { if (a.length >= 1) at.shear(Math.tan(Math.toRadians(a[0])), 0); }
                case "skewY" -> { if (a.length >= 1) at.shear(0, Math.tan(Math.toRadians(a[0]))); }
                default -> { }
            }
        }
        return any ? at : null;
    }

    // ------------------------------------------------------------ パス

    /** SVG のパスのデータ (d) を Path2D にする。 */
    static final class SvgPath {
        private final String s;
        private int i;

        private SvgPath(String s) { this.s = s; }

        static Path2D parse(String d) {
            Path2D p = new Path2D.Double();
            if (d == null || d.isBlank()) return null;
            try {
                new SvgPath(d).build(p);
            } catch (RuntimeException e) {
                // 読めたところまでを使う (SVG の仕様と同じ)
            }
            return p;
        }

        private void skip() {
            while (i < s.length() && (Character.isWhitespace(s.charAt(i)) || s.charAt(i) == ',')) i++;
        }

        private double n() {
            skip();
            int st = i;
            if (i < s.length() && (s.charAt(i) == '-' || s.charAt(i) == '+')) i++;
            boolean dot = false;
            while (i < s.length()) {
                char c = s.charAt(i);
                if (Character.isDigit(c)) i++;
                else if (c == '.' && !dot) { dot = true; i++; }
                else break;
            }
            if (i < s.length() && (s.charAt(i) == 'e' || s.charAt(i) == 'E')) {
                int save = i;
                i++;
                if (i < s.length() && (s.charAt(i) == '-' || s.charAt(i) == '+')) i++;
                if (i < s.length() && Character.isDigit(s.charAt(i))) { while (i < s.length() && Character.isDigit(s.charAt(i))) i++; }
                else i = save;
            }
            if (st == i) throw new IllegalArgumentException("number");
            return Double.parseDouble(s.substring(st, i));
        }

        /** 円弧のフラグ (0 か 1。区切りなしで続けて書けるので 1 文字だけ読む)。 */
        private boolean flag() {
            skip();
            char c = s.charAt(i++);
            if (c != '0' && c != '1') throw new IllegalArgumentException("flag");
            return c == '1';
        }

        private void build(Path2D p) {
            double cx = 0, cy = 0, sx = 0, sy = 0, lcx = 0, lcy = 0;
            char prev = ' ';
            char cmd = ' ';
            boolean started = false;
            while (true) {
                skip();
                if (i >= s.length()) break;
                char c = s.charAt(i);
                if (Character.isLetter(c) && c != 'e' && c != 'E') { cmd = c; i++; }
                else if (cmd == ' ') throw new IllegalArgumentException("cmd");
                else if (cmd == 'M') cmd = 'L';
                else if (cmd == 'm') cmd = 'l';
                boolean rel = Character.isLowerCase(cmd);
                double ox = rel ? cx : 0, oy = rel ? cy : 0;
                switch (Character.toUpperCase(cmd)) {
                    case 'M' -> { cx = ox + n(); cy = oy + n(); p.moveTo(cx, cy); sx = cx; sy = cy; started = true; }
                    case 'L' -> { cx = ox + n(); cy = oy + n(); line(p, cx, cy, started); started = true; }
                    case 'H' -> { cx = (rel ? cx : 0) + n(); line(p, cx, cy, started); started = true; }
                    case 'V' -> { cy = (rel ? cy : 0) + n(); line(p, cx, cy, started); started = true; }
                    case 'C' -> {
                        double x1 = ox + n(), y1 = oy + n(), x2 = ox + n(), y2 = oy + n();
                        cx = ox + n(); cy = oy + n();
                        p.curveTo(x1, y1, x2, y2, cx, cy);
                        lcx = x2; lcy = y2;
                    }
                    case 'S' -> {
                        boolean cubic = "CcSs".indexOf(prev) >= 0;
                        double x1 = cubic ? 2 * cx - lcx : cx, y1 = cubic ? 2 * cy - lcy : cy;
                        double x2 = ox + n(), y2 = oy + n();
                        cx = ox + n(); cy = oy + n();
                        p.curveTo(x1, y1, x2, y2, cx, cy);
                        lcx = x2; lcy = y2;
                    }
                    case 'Q' -> {
                        double x1 = ox + n(), y1 = oy + n();
                        cx = ox + n(); cy = oy + n();
                        p.quadTo(x1, y1, cx, cy);
                        lcx = x1; lcy = y1;
                    }
                    case 'T' -> {
                        boolean quad = "QqTt".indexOf(prev) >= 0;
                        double x1 = quad ? 2 * cx - lcx : cx, y1 = quad ? 2 * cy - lcy : cy;
                        cx = ox + n(); cy = oy + n();
                        p.quadTo(x1, y1, cx, cy);
                        lcx = x1; lcy = y1;
                    }
                    case 'A' -> {
                        double rx = Math.abs(n()), ry = Math.abs(n()), rot = n();
                        boolean large = flag(), sweep = flag();
                        double x = ox + n(), y = oy + n();
                        arc(p, cx, cy, rx, ry, rot, large, sweep, x, y);
                        cx = x; cy = y;
                    }
                    case 'Z' -> { p.closePath(); cx = sx; cy = sy; prev = cmd; cmd = ' '; continue; }
                    default -> throw new IllegalArgumentException("cmd " + cmd);
                }
                prev = cmd;
            }
        }

        private static void line(Path2D p, double x, double y, boolean started) {
            if (started) p.lineTo(x, y); else p.moveTo(x, y);
        }

        /** 端点で書いた円弧 (SVG 1.1 付録 F.6) を中心の形に直して加える。 */
        private static void arc(Path2D p, double x0, double y0, double rx, double ry, double angle, boolean large, boolean sweep, double x, double y) {
            if (x0 == x && y0 == y) return;
            if (rx == 0 || ry == 0) { p.lineTo(x, y); return; }
            double a = Math.toRadians(angle % 360), cos = Math.cos(a), sin = Math.sin(a);
            double dx2 = (x0 - x) / 2, dy2 = (y0 - y) / 2;
            double x1 = cos * dx2 + sin * dy2, y1 = -sin * dx2 + cos * dy2;
            double rx2 = rx * rx, ry2 = ry * ry, x12 = x1 * x1, y12 = y1 * y1;
            double check = x12 / rx2 + y12 / ry2;
            if (check > 1) { rx *= Math.sqrt(check); ry *= Math.sqrt(check); rx2 = rx * rx; ry2 = ry * ry; }
            double sign = large == sweep ? -1 : 1;
            double sq = (rx2 * ry2 - rx2 * y12 - ry2 * x12) / (rx2 * y12 + ry2 * x12);
            double coef = sign * Math.sqrt(Math.max(0, sq));
            double cx1 = coef * (rx * y1 / ry), cy1 = coef * -(ry * x1 / rx);
            double cx = (x0 + x) / 2 + (cos * cx1 - sin * cy1), cy = (y0 + y) / 2 + (sin * cx1 + cos * cy1);
            double ux = (x1 - cx1) / rx, uy = (y1 - cy1) / ry, vx = (-x1 - cx1) / rx, vy = (-y1 - cy1) / ry;
            double start = Math.toDegrees(Math.atan2(uy, ux));
            double extent = Math.toDegrees(Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy));
            if (!sweep && extent > 0) extent -= 360;
            else if (sweep && extent < 0) extent += 360;
            // Arc2D は y 上向きの角度なので符号を反転して作り、回転を掛ける
            Arc2D arc = new Arc2D.Double(cx - rx, cy - ry, 2 * rx, 2 * ry, -start, -extent, Arc2D.OPEN);
            Shape s = AffineTransform.getRotateInstance(a, cx, cy).createTransformedShape(arc);
            p.append(s.getPathIterator(null), true);
            Point2D cur = p.getCurrentPoint();
            if (cur != null && (Math.abs(cur.getX() - x) > 1e-6 || Math.abs(cur.getY() - y) > 1e-6)) p.lineTo(x, y);
        }
    }
}
