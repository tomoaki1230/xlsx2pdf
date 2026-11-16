package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.AffineTransform;
import java.awt.geom.FlatteningPathIterator;
import java.awt.geom.Path2D;
import java.awt.geom.PathIterator;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.sl.draw.geom.AdjustPoint;
import org.apache.poi.sl.draw.geom.AdjustValue;
import org.apache.poi.sl.draw.geom.ArcToCommand;
import org.apache.poi.sl.draw.geom.ClosePathCommand;
import org.apache.poi.sl.draw.geom.Context;
import org.apache.poi.sl.draw.geom.CurveToCommand;
import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.draw.geom.Guide;
import org.apache.poi.sl.draw.geom.GuideIf;
import org.apache.poi.sl.draw.geom.LineToCommand;
import org.apache.poi.sl.draw.geom.MoveToCommand;
import org.apache.poi.sl.draw.geom.PathIf;
import org.apache.poi.sl.draw.geom.QuadToCommand;
import org.w3c.dom.Element;

/**
 * ワードアートの文字の変形 (bodyPr/prstTxWarp)。形は ECMA-376 第 5 版 Part 1 の付属ファイル presetTextWarpDefinitions.xml
 * (resources の xlsx2pdf/presetTextWarpDefinitions.xml) の式で、調整値 (avLst) を使って求める。
 * パスが 2 本以上なら 1 本目 (上) と最後 (下) の間に文字の高さを合わせ、1 本なら文字をパスに沿わせる (円弧・円。
 * 文字の高さは中心までの距離の半分。パスが中心より上なら文字の上端、下なら下端をパスに置く)。
 * 字形の輪郭を細かい折れ線にしてから、点ごとに写す。
 */
final class TextWarp {
    private TextWarp() {}

    /** 1 つのラン: 字形の輪郭 (変形前、文字のまとまりの座標) と書式。 */
    private record Glyphs(Path2D path, Color fill, DrawingPainter.LineStyle line) {}

    /** 定義 (名前 → 形)。最初に使うときに読む。 */
    private static Map<String, CustomGeometry> defs;

    /** 1 本のパスに沿わせるときの文字の高さ (中心までの距離に対する割合)。 */
    static final double ALONG_HEIGHT = 0.5;

    /**
     * 定義を読む。システムプロパティ xlsx2pdf.textWarpDefinitions (ファイルの場所) があればそれを、なければ resources の
     * xlsx2pdf/presetTextWarpDefinitions.xml を読む。どちらもなければ空 (近似の形で描く)。定義のファイルは Ecma の著作物のため
     * 公開版には含めない。
     */
    private static synchronized Map<String, CustomGeometry> defs() {
        if (defs != null) return defs;
        defs = new HashMap<>();
        String file = System.getProperty("xlsx2pdf.textWarpDefinitions");
        try (InputStream in = file != null ? new java.io.FileInputStream(file)
                : TextWarp.class.getResourceAsStream("/xlsx2pdf/presetTextWarpDefinitions.xml")) {
            if (in == null) return defs;
            Element root = Xml.parse(in);
            for (Element e : Xml.children(root)) defs.put(e.getLocalName(), geometry(e));
        } catch (Exception e) {
            Warnings.add("ワードアートの変形の定義を読めませんでした", e);
        }
        return defs;
    }

    /** 定義の要素から形を組み立てる。 */
    private static CustomGeometry geometry(Element e) {
        CustomGeometry g = new CustomGeometry();
        for (Element gd : Xml.children(Xml.child(e, "avLst"), "gd")) {
            AdjustValue a = new AdjustValue();
            a.setName(gd.getAttribute("name"));
            a.setFmla(gd.getAttribute("fmla"));
            g.addAdjustGuide(a);
        }
        for (Element gd : Xml.children(Xml.child(e, "gdLst"), "gd")) {
            Guide a = new Guide();
            a.setName(gd.getAttribute("name"));
            a.setFmla(gd.getAttribute("fmla"));
            g.addGeomGuide(a);
        }
        for (Element pe : Xml.children(Xml.child(e, "pathLst"), "path")) {
            org.apache.poi.sl.draw.geom.Path p = new org.apache.poi.sl.draw.geom.Path();
            for (Element c : Xml.children(pe)) {
                List<Element> pts = Xml.children(c, "pt");
                switch (c.getLocalName()) {
                    case "moveTo" -> { MoveToCommand m = new MoveToCommand(); m.setPt(pt(pts.get(0))); p.addCommand(m); }
                    case "lnTo" -> { LineToCommand m = new LineToCommand(); m.setPt(pt(pts.get(0))); p.addCommand(m); }
                    case "arcTo" -> {
                        ArcToCommand m = new ArcToCommand();
                        m.setWR(c.getAttribute("wR")); m.setHR(c.getAttribute("hR"));
                        m.setStAng(c.getAttribute("stAng")); m.setSwAng(c.getAttribute("swAng"));
                        p.addCommand(m);
                    }
                    case "cubicBezTo" -> {
                        CurveToCommand m = new CurveToCommand();
                        m.setPt1(pt(pts.get(0))); m.setPt2(pt(pts.get(1))); m.setPt3(pt(pts.get(2)));
                        p.addCommand(m);
                    }
                    case "quadBezTo" -> { QuadToCommand m = new QuadToCommand(); m.setPt1(pt(pts.get(0))); m.setPt2(pt(pts.get(1))); p.addCommand(m); }
                    case "close" -> p.addCommand(new ClosePathCommand());
                    default -> { }
                }
            }
            g.addPath(p);
        }
        return g;
    }

    private static AdjustPoint pt(Element e) {
        AdjustPoint p = new AdjustPoint();
        p.setX(e.getAttribute("x"));
        p.setY(e.getAttribute("y"));
        return p;
    }

    /** 変形して描けるか (定義を知っているか、近似の形があるか)。 */
    static boolean known(String prst) {
        return defs().containsKey(prst) || prst.startsWith("textArch") || prst.startsWith("textCircle") || prst.startsWith("textButton")
                || ENVELOPES.contains(prst.replace("Pour", ""));
    }

    // ------------------------------------------------------------ 定義がないときの近似の形

    private static final List<String> ENVELOPES = List.of("textPlain", "textInflate", "textDeflate", "textInflateTop", "textDeflateTop",
            "textInflateBottom", "textDeflateBottom", "textWave1", "textWave2", "textDoubleWave1", "textWave4", "textTriangle",
            "textTriangleInverted", "textChevron", "textChevronInverted", "textCurveUp", "textCurveDown", "textCanUp", "textCanDown",
            "textSlantUp", "textSlantDown", "textFadeRight", "textFadeLeft", "textFadeUp", "textFadeDown", "textCascadeUp",
            "textCascadeDown", "textDeflateInflate", "textDeflateInflateDeflate", "textStop", "textRingInside", "textRingOutside");



    /**
     * 段落の文字を変形して矩形 box に描く。
     *
     * @param warpEl prstTxWarp の要素 (調整値 avLst)
     * @param line   文字の輪郭の線 (なければ null)
     */
    static void paint(PdfCanvas cv, RenderContext ctx, List<List<TextLayout.Run>> lines, Rectangle2D box, String prst, Element warpEl,
                      DrawingPainter.LineStyle line, DrawingPainter dp) throws IOException {
        paint(cv, ctx, lines, box, prst, warpEl, line, dp, null);
    }

    /** @param shadow 文字の外側の影 (最初のランの rPr の effectLst の outerShdw。なければ null) */
    static void paint(PdfCanvas cv, RenderContext ctx, List<List<TextLayout.Run>> lines, Rectangle2D box, String prst, Element warpEl,
                      DrawingPainter.LineStyle line, DrawingPainter dp, Element shadow) throws IOException {
        CustomGeometry geom = defs().get(prst);
        // 字形を並べる (行は上から順に、行の高さはフォントの大きさ × 1.2)
        List<Glyphs> gs = new ArrayList<>();
        double y = 0;
        for (List<TextLayout.Run> runs : lines) {
            double x = 0, size = 0;
            for (TextLayout.Run r : runs) size = Math.max(size, r.size);
            if (size == 0) size = 11;
            double base = y + size * 0.9;
            for (TextLayout.Run r : runs) {
                Path2D p = new Path2D.Double();
                double[] adv = new double[1];
                for (int i = 0; i < r.text.length(); ) {
                    int cp = r.text.codePointAt(i);
                    i += Character.charCount(cp);
                    if (cp == '\n' || cp == '\r') continue;
                    Path2D g = r.spec.outline(cp, r.size, ctx.fonts, adv);
                    if (g != null) p.append(g.createTransformedShape(AffineTransform.getTranslateInstance(x, base)), false);
                    x += adv[0] + r.spacing;
                }
                gs.add(new Glyphs(p, r.color, line));
            }
            y += size * 1.2;
        }
        Rectangle2D src = null;
        for (Glyphs g : gs) {
            Rectangle2D b = g.path.getBounds2D();
            if (b.isEmpty()) continue;
            src = src == null ? b : src.createUnion(b);
        }
        if (src == null) return;
        if (geom == null) {
            // 定義がない: 近似の形 (円弧は半径の半分の高さ、ほかは上下の線の間)
            String kind = prst.replace("Pour", "");
            for (Glyphs g : gs) {
                Path2D out = mapApprox(g.path, src, box, kind);
                if (shadow != null && dp != null) dp.textShadow(cv, shadow, out);
                if (g.fill != null) cv.fill(out, g.fill);
                if (g.line != null && dp != null) dp.strokeShape(cv, out, g.line);
            }
            return;
        }
        // 形のパスを求める ((0, 0, w, h) で求めて矩形の位置へ動かす)
        Map<String, GuideIf> adj = new HashMap<>();
        for (Element gd : Xml.children(Xml.child(warpEl, "avLst"), "gd")) {
            Guide g = new Guide();
            g.setName(gd.getAttribute("name"));
            g.setFmla(gd.getAttribute("fmla"));
            adj.put(g.getName(), g);
        }
        Rectangle2D anchor = new Rectangle2D.Double(0, 0, Math.max(box.getWidth(), 1e-3), Math.max(box.getHeight(), 1e-3));
        Context c = new Context(geom, anchor, adj::get);
        AffineTransform at = AffineTransform.getTranslateInstance(box.getX(), box.getY());
        List<Poly> polys = new ArrayList<>();
        for (PathIf p : geom) {
            Poly q = new Poly(new Path2D.Double(p.getPath(c), at));
            if (q.len > 1e-6) polys.add(q);
        }
        if (polys.isEmpty()) return;
        final Rectangle2D s = src;
        java.util.function.BiFunction<Double, Double, double[]> map;
        if (polys.size() >= 2) {
            Poly top = polys.get(0), bot = polys.get(polys.size() - 1);
            map = (u, v) -> {
                double[] a = top.at(u), b = bot.at(u);
                return new double[] {a[0] + v * (b[0] - a[0]), a[1] + v * (b[1] - a[1])};
            };
        } else {
            Poly p = polys.get(0);
            double cx = box.getCenterX(), cy = box.getCenterY();
            double[] mid = p.at(0.5);
            boolean topOnPath = mid[1] <= cy;
            map = (u, v) -> {
                double[] a = p.at(u);
                double k = (topOnPath ? v : 1 - v) * ALONG_HEIGHT;
                return new double[] {a[0] + k * (cx - a[0]), a[1] + k * (cy - a[1])};
            };
        }
        for (Glyphs g : gs) {
            Path2D out = new Path2D.Double(g.path.getWindingRule());
            double[] q = new double[6];
            for (PathIterator it = new FlatteningPathIterator(g.path.getPathIterator(null), 0.05); !it.isDone(); it.next()) {
                int t = it.currentSegment(q);
                if (t == PathIterator.SEG_CLOSE) { out.closePath(); continue; }
                double u = (q[0] - s.getX()) / s.getWidth(), v = (q[1] - s.getY()) / s.getHeight();
                double[] m = map.apply(u, v);
                if (t == PathIterator.SEG_MOVETO) out.moveTo(m[0], m[1]); else out.lineTo(m[0], m[1]);
            }
            if (shadow != null && dp != null) dp.textShadow(cv, shadow, out);
            if (g.fill != null) cv.fill(out, g.fill);
            if (g.line != null && dp != null) dp.strokeShape(cv, out, g.line);
        }
    }

    /** 折れ線にしたパス (最初の部分パスだけ) と、長さの割合での位置。 */
    private static final class Poly {
        final List<double[]> pts = new ArrayList<>();
        final List<Double> acc = new ArrayList<>();
        double len;

        Poly(Path2D p) {
            double[] q = new double[6];
            for (PathIterator it = new FlatteningPathIterator(p.getPathIterator(null), 0.05); !it.isDone(); it.next()) {
                int t = it.currentSegment(q);
                if (t == PathIterator.SEG_MOVETO && !pts.isEmpty()) break;
                if (t == PathIterator.SEG_CLOSE) continue;
                if (!Double.isFinite(q[0]) || !Double.isFinite(q[1])) continue;
                if (!pts.isEmpty()) {
                    double[] l = pts.get(pts.size() - 1);
                    len += Math.hypot(q[0] - l[0], q[1] - l[1]);
                }
                pts.add(new double[] {q[0], q[1]});
                acc.add(len);
            }
        }

        /** 長さの割合 u (0〜1) の位置。 */
        double[] at(double u) {
            if (pts.size() == 1) return pts.get(0);
            double d = Math.max(0, Math.min(1, u)) * len;
            int lo = 0, hi = acc.size() - 1;
            while (hi - lo > 1) {
                int m = (lo + hi) / 2;
                if (acc.get(m) <= d) lo = m; else hi = m;
            }
            double a = acc.get(lo), b = acc.get(hi);
            double f = b > a ? (d - a) / (b - a) : 0;
            double[] p = pts.get(lo), q = pts.get(hi);
            return new double[] {p[0] + f * (q[0] - p[0]), p[1] + f * (q[1] - p[1])};
        }
    }

    /** 近似の形: 輪郭を細かい折れ線にしてから、点ごとに写す。 */
    private static Path2D mapApprox(Path2D p, Rectangle2D src, Rectangle2D box, String kind) {
        Path2D out = new Path2D.Double(p.getWindingRule());
        double[] c = new double[6];
        for (PathIterator it = new FlatteningPathIterator(p.getPathIterator(null), 0.05); !it.isDone(); it.next()) {
            int t = it.currentSegment(c);
            if (t == PathIterator.SEG_CLOSE) { out.closePath(); continue; }
            double u = (c[0] - src.getX()) / src.getWidth(), v = (c[1] - src.getY()) / src.getHeight();
            double[] q = point(kind, u, v, box);
            if (t == PathIterator.SEG_MOVETO) out.moveTo(q[0], q[1]); else out.lineTo(q[0], q[1]);
        }
        return out;
    }

    /** 文字のまとまりの中の位置 (u: 左 0〜右 1、v: 上 0〜下 1) を、変形後の位置にする。 */
    private static double[] point(String kind, double u, double v, Rectangle2D b) {
        double w = b.getWidth(), h = b.getHeight(), cx = b.getCenterX(), cy = b.getCenterY();
        switch (kind) {
            case "textArchUp", "textArchDown", "textCircle", "textButton" -> {
                // 円弧に沿わせる: 文字の高さは半径の 50%。上の円弧は左から右へ、下の円弧は文字の上を中心に向ける
                double th = 0.5;
                double ang, r;
                if (kind.equals("textArchDown")) { ang = Math.PI + u * Math.PI; r = 1 - th + v * th; }
                else if (kind.equals("textCircle")) { ang = Math.PI - u * 2 * Math.PI; r = 1 - v * th; }
                else { ang = Math.PI - u * Math.PI; r = 1 - v * th; }
                return new double[] {cx + w / 2 * r * Math.cos(ang), cy - h / 2 * r * Math.sin(ang)};
            }
            default -> {
                double[] tb = envelope(kind, u);
                double top = b.getY() + tb[0] * h, bot = b.getY() + tb[1] * h;
                double x = b.getX() + u * w;
                if (kind.equals("textFadeUp") || kind.equals("textFadeDown")) {
                    // 上 (下) を狭める
                    double f = kind.equals("textFadeUp") ? 1 - v : v;
                    double shrink = 0.25 * f;
                    x = b.getX() + w * (shrink + u * (1 - 2 * shrink));
                }
                return new double[] {x, top + v * (bot - top)};
            }
        }
    }

    /** 包絡線の上と下 (0〜1、上が 0)。 */
    private static double[] envelope(String kind, double u) {
        double s = Math.sin(Math.PI * u), a = 0.2, m = 1 - Math.abs(2 * u - 1);
        return switch (kind) {
            case "textInflate" -> new double[] {a * (1 - s), 1 - a * (1 - s)};
            case "textDeflate" -> new double[] {a * s, 1 - a * s};
            case "textInflateTop" -> new double[] {a * 1.5 * (1 - s), 1};
            case "textDeflateTop" -> new double[] {a * 1.5 * s, 1};
            case "textInflateBottom" -> new double[] {0, 1 - a * 1.5 * (1 - s)};
            case "textDeflateBottom" -> new double[] {0, 1 - a * 1.5 * s};
            case "textDeflateInflate" -> new double[] {a * s, 1 - a * (1 - s)};
            case "textDeflateInflateDeflate" -> new double[] {a * Math.abs(Math.sin(2 * Math.PI * u)), 1 - a * Math.abs(Math.sin(2 * Math.PI * u))};
            case "textWave1" -> { double d = 0.1 * Math.sin(2 * Math.PI * u); yield new double[] {0.1 - d, 0.9 - d}; }
            case "textWave2" -> { double d = 0.1 * Math.sin(2 * Math.PI * u); yield new double[] {0.1 + d, 0.9 + d}; }
            case "textDoubleWave1", "textWave4" -> { double d = 0.1 * Math.sin(4 * Math.PI * u); yield new double[] {0.1 - d, 0.9 - d}; }
            case "textTriangle" -> new double[] {0.9 * (1 - m), 1};
            case "textTriangleInverted" -> new double[] {0, 1 - 0.9 * (1 - m)};
            case "textChevron" -> new double[] {0.3 * (1 - m), 1 - 0.3 * m};
            case "textChevronInverted" -> new double[] {0.3 * m, 1 - 0.3 * (1 - m)};
            case "textCurveUp" -> new double[] {0.3 * (1 - u) * (1 - u), 1 - 0.3 * u * u};
            case "textCurveDown" -> new double[] {0.3 * u * u, 1 - 0.3 * (1 - u) * (1 - u)};
            case "textCanUp" -> new double[] {a * s, 1 - a * (1 - s) - a};
            case "textCanDown" -> new double[] {a * (1 - s) + 0, 1 - a * s};
            case "textSlantUp" -> new double[] {0.3 * (1 - u), 1 - 0.3 * u};
            case "textSlantDown" -> new double[] {0.3 * u, 1 - 0.3 * (1 - u)};
            case "textFadeRight" -> new double[] {0.25 * u, 1 - 0.25 * u};
            case "textFadeLeft" -> new double[] {0.25 * (1 - u), 1 - 0.25 * (1 - u)};
            case "textCascadeUp" -> new double[] {0.5 * (1 - u), 1 - 0.2 * u};
            case "textCascadeDown" -> new double[] {0.5 * u, 1 - 0.2 * (1 - u)};
            case "textStop" -> { double e = Math.min(1, Math.min(u, 1 - u) * 5); yield new double[] {0.2 * (1 - e), 1 - 0.2 * (1 - e)}; }
            case "textRingInside", "textRingOutside" -> new double[] {0.2 * (1 - s), 1 - 0.2 * (1 - s)};
            default -> new double[] {0, 1}; // textPlain
        };
    }
}
