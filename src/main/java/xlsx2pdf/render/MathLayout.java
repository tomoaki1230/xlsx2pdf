package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.AffineTransform;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.Element;

/**
 * 数式 (OMML: m:oMathPara / m:oMath) の組版。数式の部品を、幅・基線より上の高さ・下の深さを持つ箱に組み立て、
 * Cambria Math の字形で描く。上付き・下付き・分数・根号・n 項演算子 (総和・積分)・かっこ・行列・アクセント・
 * 上線・関数・極限・数式の配列に対応する。寸法は TeX と Cambria Math の MATH 表の値に近い近似。
 */
final class MathLayout {
    /** 表示の大きさの段階: 0 ディスプレイ、1 テキスト、2 スクリプト、3 スクリプトのスクリプト。 */
    private static final int D = 0, T = 1, S = 2, SS = 3;
    /** スクリプトの大きさ (元の文字の倍)。 */
    static final double SCRIPT = 0.73, SCRIPT2 = 0.6;
    /** 数式の軸の高さ (分数の線・かっこの中心。em)。 */
    static final double AXIS = 0.25;
    /** 分数の線の太さ、根号の線の太さ (em)。 */
    static final double RULE = 0.055, RAD_RULE = 0.05;
    /** 文字の箱の基線より上の高さと下の深さ (em)。 */
    static final double TEXT_ASC = 0.72, TEXT_DSC = 0.22;
    /** 2 項演算子・関係演算子の両側の空き (em)。 */
    static final double BIN_SPACE = 0.255, REL_SPACE = 0.26, THIN = 3.0 / 18;
    /** 数式の段落の上の空き (em。Excel の印刷の実測)。 */
    static final double PARA_TOP = 0, PARA_BOTTOM = 0.1;

    /** 数式の中の最初のラン (a:rPr の sz) の文字の大きさ。なければ dflt。 */
    static float firstSize(Element m, float dflt) {
        org.w3c.dom.NodeList l = m.getElementsByTagNameNS(A_NS, "rPr");
        for (int i = 0; i < l.getLength(); i++) {
            Element e = (Element) l.item(i);
            if (e.hasAttribute("sz")) return Xml.intAttr(e, "sz", 1100) / 100f;
        }
        return dflt;
    }

    private final RenderContext ctx;
    private final FontManager.Spec font;
    private final float baseSize;
    private final Color baseColor;

    /**
     * @param size  文字の大きさ (pt。ランに指定がなければ使う)
     * @param color 文字の色 (ランに指定がなければ使う)
     */
    MathLayout(RenderContext ctx, float size, Color color) {
        this.ctx = ctx;
        this.font = ctx.fonts.spec("Cambria Math", false, false);
        this.baseSize = size;
        this.baseColor = color == null ? Color.BLACK : color;
    }

    // ------------------------------------------------------------ 箱

    /** 箱: 幅、基線より上の高さ、下の深さ (pt)。描画は (x, 基線)。 */
    abstract static class Box {
        double w, asc, dsc;
        /** 種類 (演算子の空きの決定): 0 普通、1 2 項演算子、2 関係、3 開きかっこ、4 閉じかっこ、5 句読点、6 大きな演算子。 */
        int kind;

        abstract void draw(PdfCanvas cv, double x, double base) throws IOException;
    }

    /** 文字列。 */
    static final class Text extends Box {
        final String s;
        final FontManager.Spec spec;
        final float size;
        final Color color;

        Text(String s, FontManager.Spec spec, float size, Color color, FontManager fm) {
            this.s = s;
            this.spec = spec;
            this.size = size;
            this.color = color;
            w = spec.width(s, size, fm);
            // 字形の実際の高さ (Cambria Math の ascent・descent は大きな記号のための値で、文字には大きすぎる)
            asc = TEXT_ASC * size;
            dsc = TEXT_DSC * size;
        }

        @Override
        void draw(PdfCanvas cv, double x, double base) throws IOException {
            cv.text(s, x, base, spec, size, color, 1, 0);
        }
    }

    /** 子の箱を並べた箱 (子は (dx, dy) に置く。dy は下向き)。 */
    static final class Group extends Box {
        final List<Box> kids = new ArrayList<>();
        final List<double[]> pos = new ArrayList<>();
        final List<Runnable> none = List.of();
        /** 線・図形 (x, 基線からの相対座標の形、塗るか、線の太さ、色)。 */
        final List<Object[]> shapes = new ArrayList<>();

        void add(Box b, double dx, double dy) {
            kids.add(b);
            pos.add(new double[] {dx, dy});
            w = Math.max(w, dx + b.w);
            asc = Math.max(asc, b.asc - dy);
            dsc = Math.max(dsc, b.dsc + dy);
        }

        void shape(java.awt.Shape s, boolean fill, double lw, Color c) {
            shapes.add(new Object[] {s, fill, lw, c});
            Rectangle2D r = s.getBounds2D();
            w = Math.max(w, r.getMaxX());
            asc = Math.max(asc, -r.getY());
            dsc = Math.max(dsc, r.getMaxY());
        }

        @Override
        void draw(PdfCanvas cv, double x, double base) throws IOException {
            for (int i = 0; i < kids.size(); i++) kids.get(i).draw(cv, x + pos.get(i)[0], base + pos.get(i)[1]);
            for (Object[] o : shapes) {
                java.awt.Shape s = AffineTransform.getTranslateInstance(x, base).createTransformedShape((java.awt.Shape) o[0]);
                if ((Boolean) o[1]) cv.fill(s, (Color) o[3]);
                else cv.stroke(s, new PdfCanvas.Stroke((float) (double) (Double) o[2]), (Color) o[3]);
            }
        }
    }

    // ------------------------------------------------------------ 組み立て

    /** 文字の大きさ (段階 st の em)。 */
    private static float scaled(float size, int st) {
        return (float) (st <= T ? size : st == S ? size * SCRIPT : size * SCRIPT2);
    }

    private static int script(int st) { return st <= T ? S : SS; }

    /** 段落の中の数式 (a14:m の中身: oMathPara または oMath) の箱。複数の行はディスプレイで縦に並べる。 */
    Box build(Element m) {
        Element para = Xml.child(m, "oMathPara");
        if (para != null) {
            List<Box> lines = new ArrayList<>();
            for (Element om : Xml.children(para, "oMath")) lines.add(row(om, D, baseSize, baseColor));
            return stack(lines, 0.3 * baseSize, 0);
        }
        Element om = Xml.child(m, "oMath");
        return row(om != null ? om : m, T, baseSize, baseColor);
    }

    /** oMathPara の配置 (left, right, center, centerGroup)。 */
    static String justification(Element m) {
        Element para = Xml.child(m, "oMathPara");
        return mv(Xml.child(Xml.child(para, "oMathParaPr"), "jc"), para != null ? "centerGroup" : "inline");
    }

    /** 縦に並べる (行の間 gap、align: 0 中央、-1 左)。行の中心を軸に合わせず、最初の行の基線を基線にする。 */
    private Box stack(List<Box> lines, double gap, int align) {
        if (lines.size() == 1) return lines.get(0);
        double w = 0;
        for (Box b : lines) w = Math.max(w, b.w);
        Group g = new Group();
        double y = 0;
        for (int i = 0; i < lines.size(); i++) {
            Box b = lines.get(i);
            if (i > 0) y += lines.get(i - 1).dsc + gap + b.asc;
            g.add(b, align == 0 ? (w - b.w) / 2 : 0, y);
        }
        return g;
    }

    /** 要素の並び (oMath, e, num, sub など) を 1 行に組む。演算子の両側に空きを入れる。 */
    private Box row(Element parent, int st, float size, Color color) {
        List<Box> items = new ArrayList<>();
        if (parent != null) for (Element c : Xml.children(parent)) element(c, st, size, color, items);
        Group g = new Group();
        double x = 0, em = scaled(size, st);
        int prev = -1;
        for (int i = 0; i < items.size(); i++) {
            Box b = items.get(i);
            int k = b.kind;
            // 2 項演算子は、前が普通の要素・閉じかっこのときだけ 2 項 (それ以外は単項で空きなし)
            if (k == 1 && (prev == -1 || prev == 1 || prev == 2 || prev == 3 || prev == 5)) k = 0;
            double sp = 0;
            if (st <= T) {
                if (k == 1) sp = BIN_SPACE * em;
                else if (k == 2) sp = REL_SPACE * em;
            }
            if (i > 0 && (k == 1 || k == 2)) x += sp;
            if (i > 0 && prev >= 0 && (prev == 1 || prev == 2) && b.kind != 1 && b.kind != 2) x += prevSpace(prev, st, em, items, i);
            if (prev == 5 && st <= T) x += THIN * em;
            g.add(b, x, 0);
            x += b.w;
            prev = k;
        }
        if (items.isEmpty()) { g.asc = em * 0.6; g.w = 0; }
        return g;
    }

    /** 前の演算子の後ろの空き。 */
    private double prevSpace(int prev, int st, double em, List<Box> items, int i) {
        if (st > T) return 0;
        // 前の演算子が単項に変わっていれば空けない
        if (i >= 2) {
            int pp = items.get(i - 2).kind;
            if (prev == 1 && (pp == 1 || pp == 2 || pp == 3 || pp == 5)) return 0;
        } else if (prev == 1) return 0;
        return prev == 1 ? BIN_SPACE * em : REL_SPACE * em;
    }

    /** 1 つの要素を組んで items に加える。 */
    private void element(Element c, int st, float size, Color color, List<Box> items) {
        String n = c.getLocalName();
        float em = scaled(size, st);
        switch (n) {
            case "r" -> run(c, st, size, color, items);
            case "sSup", "sSub", "sSubSup" -> {
                Box base = row(Xml.child(c, "e"), st, size, color);
                Box sub = n.equals("sSup") ? null : row(Xml.child(c, "sub"), script(st), size, color);
                Box sup = n.equals("sSub") ? null : row(Xml.child(c, "sup"), script(st), size, color);
                items.add(scripts(base, sub, sup, em));
            }
            case "sPre" -> {
                Box base = row(Xml.child(c, "e"), st, size, color);
                Box sub = row(Xml.child(c, "sub"), script(st), size, color), sup = row(Xml.child(c, "sup"), script(st), size, color);
                Box s = scripts(new Group(), sub, sup, em);
                Group g = new Group();
                g.add(s, 0, 0);
                g.add(base, s.w, 0);
                items.add(g);
            }
            case "f" -> items.add(fraction(c, st, size, color));
            case "rad" -> items.add(radical(c, st, size, color));
            case "nary" -> items.add(nary(c, st, size, color));
            case "d" -> items.add(delimited(c, st, size, color));
            case "m" -> items.add(matrix(c, st, size, color));
            case "acc" -> items.add(accent(c, st, size, color));
            case "bar" -> {
                Box b = row(Xml.child(c, "e"), st, size, color);
                boolean top = "top".equals(mv(Xml.child(Xml.child(c, "barPr"), "pos"), "bot"));
                Group g = new Group();
                g.add(b, 0, 0);
                double t = RULE * em;
                double y = top ? -b.asc - 0.12 * em : b.dsc + 0.12 * em;
                g.shape(new Rectangle2D.Double(0, y - t / 2, b.w, t), true, 0, color);
                items.add(g);
            }
            case "func" -> {
                Box name = row(Xml.child(c, "fName"), st, size, color);
                Box arg = row(Xml.child(c, "e"), st, size, color);
                Group g = new Group();
                g.add(name, 0, 0);
                g.add(arg, name.w + THIN * em, 0);
                items.add(g);
            }
            case "limLow", "limUpp" -> {
                Box b = row(Xml.child(c, "e"), st, size, color);
                Box lim = row(Xml.child(c, "lim"), script(st), size, color);
                Group g = new Group();
                double w = Math.max(b.w, lim.w);
                g.add(b, (w - b.w) / 2, 0);
                if (n.equals("limLow")) g.add(lim, (w - lim.w) / 2, b.dsc + 0.1 * em + lim.asc);
                else g.add(lim, (w - lim.w) / 2, -b.asc - 0.1 * em - lim.dsc);
                items.add(g);
            }
            case "groupChr" -> {
                Box b = row(Xml.child(c, "e"), st, size, color);
                Element pr = Xml.child(c, "groupChrPr");
                String ch = mv(Xml.child(pr, "chr"), "⏟");
                boolean top = "top".equals(mv(Xml.child(pr, "pos"), "bot"));
                Text t = new Text(ch, font, (float) (em * 0.8), color, ctx.fonts);
                Group g = new Group();
                double w = Math.max(b.w, t.w);
                g.add(b, (w - b.w) / 2, 0);
                g.add(t, (w - t.w) / 2, top ? -b.asc - t.dsc - 0.05 * em : b.dsc + t.asc + 0.05 * em);
                items.add(g);
            }
            case "eqArr" -> {
                List<Box> lines = new ArrayList<>();
                for (Element e : Xml.children(c, "e")) lines.add(row(e, st, size, color));
                Box s = stack(lines, 0.25 * em, 0);
                items.add(centerOnAxis(s, em));
            }
            case "borderBox" -> {
                Box b = row(Xml.child(c, "e"), st, size, color);
                Group g = new Group();
                double p = 0.1 * em;
                g.add(b, p, 0);
                g.shape(new Rectangle2D.Double(0, -b.asc - p, b.w + 2 * p, b.asc + b.dsc + 2 * p), false, RULE * em, color);
                items.add(g);
            }
            case "phant" -> {
                Box b = row(Xml.child(c, "e"), st, size, color);
                boolean show = mb(Xml.child(Xml.child(c, "phantPr"), "show"), true);
                if (show) items.add(b);
                else { Group g = new Group(); g.w = b.w; g.asc = b.asc; g.dsc = b.dsc; items.add(g); }
            }
            case "box", "e", "oMath" -> items.add(row(n.equals("box") ? Xml.child(c, "e") : c, st, size, color));
            default -> { } // 書式 (…Pr)・制御 (ctrlPr) などは描かない
        }
    }

    /** 上付き・下付き。 */
    private Box scripts(Box base, Box sub, Box sup, double em) {
        Group g = new Group();
        g.add(base, 0, 0);
        double x = base.w + 0.03 * em;
        double up = 0, down = 0;
        if (sup != null) up = Math.max(0.36 * em, base.asc - sup.asc * 0.5);
        if (sub != null) down = Math.max(0.24 * em, base.dsc + 0.02 * em); // Excel の印刷: 文字の下付きは基線の 0.24em 下
        if (sub != null && sup != null) {
            // 上下の間を空ける
            double gap = (up - sup.dsc) - (sub.asc - down);
            if (gap < 0.15 * em) { double d = (0.15 * em - gap) / 2; up += d; down += d; }
        }
        if (sup != null) g.add(sup, x, -up);
        if (sub != null) g.add(sub, x, down);
        g.kind = base.kind == 6 ? 6 : 0;
        return g;
    }

    /** 分数 (bar、lin、skw、noBar)。 */
    private Box fraction(Element c, int st, float size, Color color) {
        String type = mv(Xml.child(Xml.child(c, "fPr"), "type"), "bar");
        double em = scaled(size, st);
        if (type.equals("lin") || type.equals("skw")) {
            Box num = row(Xml.child(c, "num"), st, size, color), den = row(Xml.child(c, "den"), st, size, color);
            Text sl = new Text("/", font, (float) em, color, ctx.fonts);
            Group g = new Group();
            g.add(num, 0, type.equals("skw") ? -0.25 * em : 0);
            g.add(sl, num.w, 0);
            g.add(den, num.w + sl.w, type.equals("skw") ? 0.2 * em : 0);
            return g;
        }
        int inner = st == D ? T : script(st);
        Box num = row(Xml.child(c, "num"), inner, size, color), den = row(Xml.child(c, "den"), inner, size, color);
        double w = Math.max(num.w, den.w) + 0.12 * em;
        double axis = AXIS * em, t = RULE * em, gap = (st == D ? 0.2 : 0.1) * em;
        Group g = new Group();
        g.add(num, (w - num.w) / 2, -(axis + t / 2 + gap + num.dsc));
        g.add(den, (w - den.w) / 2, -axis + t / 2 + gap + den.asc);
        if (!type.equals("noBar")) g.shape(new Rectangle2D.Double(0.05 * em, -axis - t / 2, w - 0.1 * em, t), true, 0, color);
        g.w = w + 0.05 * em;
        return g;
    }

    /** 根号 (deg が空なら平方根)。 */
    private Box radical(Element c, int st, float size, Color color) {
        double em = scaled(size, st);
        Box body = row(Xml.child(c, "e"), st, size, color);
        boolean degHide = mb(Xml.child(Xml.child(c, "radPr"), "degHide"), false);
        Element degEl = Xml.child(c, "deg");
        Box deg = degHide || degEl == null || Xml.children(degEl).isEmpty() ? null : row(degEl, SS, size, color);
        double lw = RAD_RULE * em, gap = 0.1 * em;
        double top = -Math.max(body.asc, 0.62 * em) - gap - lw / 2, bottom = Math.max(body.dsc, 0.2 * em);
        double h = bottom - top;
        double sw = 0.42 * em + h * 0.06; // 根号の記号の幅
        double ox = deg != null ? Math.max(0, deg.w - sw * 0.45) : 0;
        Path2D p = new Path2D.Double();
        p.moveTo(ox, bottom - h * 0.42);
        p.lineTo(ox + sw * 0.22, bottom - h * 0.5);
        p.lineTo(ox + sw * 0.52, bottom);
        p.lineTo(ox + sw, top);
        p.lineTo(ox + sw + body.w + 0.08 * em, top);
        Group g = new Group();
        g.shape(p, false, lw, color);
        g.add(body, ox + sw + 0.04 * em, 0);
        if (deg != null) g.add(deg, 0, bottom - h * 0.55 - deg.dsc);
        g.w = ox + sw + body.w + 0.12 * em;
        return g;
    }

    /** n 項演算子 (∑、∫ など)。 */
    private Box nary(Element c, int st, float size, Color color) {
        Element pr = Xml.child(c, "naryPr");
        String ch = mv(Xml.child(pr, "chr"), "∫");
        boolean integral = "∫∬∭∮∯∰".contains(ch);
        String loc = mv(Xml.child(pr, "limLoc"), integral ? "subSup" : "undOvr");
        boolean subHide = mb(Xml.child(pr, "subHide"), false), supHide = mb(Xml.child(pr, "supHide"), false);
        double em = scaled(size, st);
        // 演算子の字形の高さ (em。Excel の印刷の実測: ディスプレイの積分 1.85、総和 1.2)
        double k = st == D ? (integral ? 1.85 : 1.2) : (integral ? 1.1 : 0.9);
        Box op = bigGlyph(ch, em, k, color);
        Box sub = subHide ? null : row(Xml.child(c, "sub"), script(st), size, color);
        Box sup = supHide ? null : row(Xml.child(c, "sup"), script(st), size, color);
        if (sub != null && sub.w == 0) sub = null;
        if (sup != null && sup.w == 0) sup = null;
        Box body = row(Xml.child(c, "e"), st, size, color);
        Group g = new Group();
        double x;
        if (loc.equals("undOvr")) {
            double w = Math.max(op.w, Math.max(sub == null ? 0 : sub.w, sup == null ? 0 : sup.w));
            g.add(op, (w - op.w) / 2, 0);
            if (sup != null) g.add(sup, (w - sup.w) / 2, -op.asc - 0.1 * em - sup.dsc);
            if (sub != null) g.add(sub, (w - sub.w) / 2, op.dsc + 0.1 * em + sub.asc);
            x = w;
        } else {
            g.add(op, 0, 0);
            double sx = op.w + (integral ? -0.05 * em : 0.03 * em);
            double sw = 0;
            if (sup != null) { g.add(sup, sx + (integral ? 0.15 * em : 0), -op.asc + sup.asc * 0.8); sw = Math.max(sw, sup.w + (integral ? 0.15 * em : 0)); }
            if (sub != null) { g.add(sub, sx - (integral ? 0.05 * em : 0), op.dsc - sub.dsc * 0.3); sw = Math.max(sw, sub.w); }
            x = sx + sw;
        }
        g.add(body, x + THIN * em, 0);
        g.kind = 6;
        return g;
    }

    /** 字形を高さ k em に拡大して軸を中心に置いた箱。 */
    private Box bigGlyph(String ch, double em, double k, Color color) {
        double[] adv = new double[1];
        Path2D p0 = font.outline(ch.codePointAt(0), (float) em, ctx.fonts, adv);
        if (p0 == null) return new Text(ch, font, (float) (em * k), color, ctx.fonts);
        double f = k * em / Math.max(1e-6, p0.getBounds2D().getHeight());
        Path2D p = new Path2D.Double(p0, AffineTransform.getScaleInstance(f, f));
        Rectangle2D b = p.getBounds2D();
        double cy = b.getCenterY(), axis = -AXIS * em;
        Path2D q = new Path2D.Double(p, AffineTransform.getTranslateInstance(-b.getX() + 0.05 * em, axis - cy));
        Group g = new Group();
        g.shape(q, true, 0, color);
        g.w = b.getWidth() + 0.1 * em;
        return g;
    }

    /** かっこで囲む (d)。中身が高ければ字形を縦に伸ばす。区切り (sepChr) にも対応。 */
    private Box delimited(Element c, int st, float size, Color color) {
        Element pr = Xml.child(c, "dPr");
        String beg = Xml.child(pr, "begChr") != null ? mv(Xml.child(pr, "begChr"), "") : "(";
        String end = Xml.child(pr, "endChr") != null ? mv(Xml.child(pr, "endChr"), "") : ")";
        String sep = Xml.child(pr, "sepChr") != null ? mv(Xml.child(pr, "sepChr"), "|") : "|";
        double em = scaled(size, st);
        List<Box> parts = new ArrayList<>();
        for (Element e : Xml.children(c, "e")) parts.add(row(e, st, size, color));
        double asc = 0.7 * em, dsc = 0.2 * em;
        for (Box b : parts) { asc = Math.max(asc, b.asc); dsc = Math.max(dsc, b.dsc); }
        // 軸を中心に上下対称の高さ
        double half = Math.max(asc - AXIS * em, dsc + AXIS * em) + 0.06 * em;
        Group g = new Group();
        double x = 0;
        if (!beg.isEmpty()) { Box d = delim(beg, half, em, color); g.add(d, x, 0); x += d.w; }
        for (int i = 0; i < parts.size(); i++) {
            if (i > 0) { Box d = delim(sep, half, em, color); g.add(d, x, 0); x += d.w; }
            g.add(parts.get(i), x, 0);
            x += parts.get(i).w;
        }
        if (!end.isEmpty()) { Box d = delim(end, half, em, color); g.add(d, x, 0); x += d.w; }
        g.w = x;
        return g;
    }

    /** かっこの字形 (高さ 2 × half に縦に伸ばす。普通の高さなら字形のまま)。 */
    private Box delim(String ch, double half, double em, Color color) {
        double[] adv = new double[1];
        Path2D p = font.outline(ch.codePointAt(0), (float) em, ctx.fonts, adv);
        if (p == null) return new Text(ch, font, (float) em, color, ctx.fonts);
        Rectangle2D b = p.getBounds2D();
        double want = 2 * half;
        double sy = Math.max(1, want / b.getHeight());
        double sx = sy > 1 ? Math.min(1.6, 1 + (sy - 1) * 0.15) : 1;
        AffineTransform at = new AffineTransform();
        at.translate(0.04 * em - b.getX() * sx, -AXIS * em);
        at.scale(sx, sy);
        at.translate(0, -b.getCenterY());
        Group g = new Group();
        g.shape(at.createTransformedShape(p), true, 0, color);
        g.w = b.getWidth() * sx + 0.08 * em;
        g.kind = "([{⟨⌊⌈|‖".contains(ch) ? 3 : 4;
        return g;
    }

    /** 行列: 列ごとに中央揃え、軸を中心に置く。 */
    private Box matrix(Element c, int st, float size, Color color) {
        double em = scaled(size, st);
        List<List<Box>> rows = new ArrayList<>();
        int nc = 0;
        for (Element mr : Xml.children(c, "mr")) {
            List<Box> r = new ArrayList<>();
            for (Element e : Xml.children(mr, "e")) r.add(row(e, st, size, color));
            rows.add(r);
            nc = Math.max(nc, r.size());
        }
        double[] cw = new double[nc];
        for (List<Box> r : rows) for (int j = 0; j < r.size(); j++) cw[j] = Math.max(cw[j], r.get(j).w);
        double colGap = 1.0 * em, rowGap = 0.25 * em;
        Group g = new Group();
        double y = 0;
        for (int i = 0; i < rows.size(); i++) {
            List<Box> r = rows.get(i);
            double ra = 0.6 * em, rd = 0.15 * em;
            for (Box b : r) { ra = Math.max(ra, b.asc); rd = Math.max(rd, b.dsc); }
            y += (i == 0 ? 0 : rowGap) + ra;
            double x = 0;
            for (int j = 0; j < r.size(); j++) {
                g.add(r.get(j), x + (cw[j] - r.get(j).w) / 2, y);
                x += cw[j] + colGap;
            }
            y += rd;
        }
        return centerOnAxis(g, em);
    }

    /** 箱の上下の中心を軸に合わせる。 */
    private static Box centerOnAxis(Box b, double em) {
        double mid = (b.dsc - b.asc) / 2; // 基線からの中心 (下向き)
        Group g = new Group();
        g.add(b, 0, -AXIS * em - mid);
        return g;
    }

    /** アクセント (既定は ^)。 */
    private Box accent(Element c, int st, float size, Color color) {
        double em = scaled(size, st);
        Box b = row(Xml.child(c, "e"), st, size, color);
        String ch = mv(Xml.child(Xml.child(c, "accPr"), "chr"), "̂");
        Group g = new Group();
        g.add(b, 0, 0);
        double top = -Math.max(b.asc, 0.45 * em);
        if (ch.equals("̅") || ch.equals("¯") || ch.equals("‾")) {
            double t = RULE * em;
            g.shape(new Rectangle2D.Double(b.w * 0.1, top - 0.1 * em - t, b.w * 0.85, t), true, 0, color);
            return g;
        }
        String sp = switch (ch) {
            case "̂" -> "ˆ";
            case "̃" -> "˜";
            case "̇" -> "˙";
            case "̈" -> "¨";
            case "́" -> "´";
            case "̀" -> "`";
            case "̌" -> "ˇ";
            case "̆" -> "˘";
            case "⃗", "⃑" -> "→";
            default -> ch;
        };
        Text a = new Text(sp, font, (float) (em * (sp.equals("→") ? 0.6 : 1)), color, ctx.fonts);
        double[] adv = new double[1];
        Path2D p = font.outline(sp.codePointAt(0), (float) (em * (sp.equals("→") ? 0.6 : 1)), ctx.fonts, adv);
        if (p != null) {
            Rectangle2D r = p.getBounds2D();
            double dx = (b.w - r.getWidth()) / 2 - r.getX() + 0.05 * em, dy = top - 0.08 * em - r.getMaxY();
            g.shape(new Path2D.Double(p, AffineTransform.getTranslateInstance(dx, dy)), true, 0, color);
        } else g.add(a, (b.w - a.w) / 2, top - 0.1 * em);
        return g;
    }

    // ------------------------------------------------------------ XML

    static final String M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math",
            A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /** 数式の属性 m:val (なければ dflt)。 */
    static String mv(Element e, String dflt) {
        if (e == null) return dflt;
        if (e.hasAttributeNS(M_NS, "val")) return e.getAttributeNS(M_NS, "val");
        if (e.hasAttribute("m:val")) return e.getAttribute("m:val");
        if (e.hasAttribute("val")) return e.getAttribute("val");
        return dflt;
    }

    /** 真偽の m:val (要素だけで値がなければ真)。 */
    static boolean mb(Element e, boolean dflt) {
        if (e == null) return dflt;
        String v = mv(e, null);
        return v == null || v.equals("1") || v.equalsIgnoreCase("on") || v.equalsIgnoreCase("true");
    }

    /** 名前空間とローカル名で子を探す。 */
    static Element childNs(Element e, String local, String ns) {
        if (e == null) return null;
        for (Element c : Xml.children(e)) if (local.equals(c.getLocalName()) && ns.equals(c.getNamespaceURI())) return c;
        return null;
    }

    // ------------------------------------------------------------ 文字

    /** ラン (m:r): 文字を数式の字形 (斜体の英字など) にし、演算子ごとに分ける。 */
    private void run(Element r, int st, float size, Color color, List<Box> items) {
        Element mrPr = childNs(r, "rPr", M_NS);
        String sty = mv(Xml.child(mrPr, "sty"), "i");
        boolean nor = mb(Xml.child(mrPr, "nor"), false);
        if (nor) sty = "p";
        Element arPr = childNs(r, "rPr", A_NS);
        float sz = size;
        Color col = color;
        if (arPr != null) {
            if (arPr.hasAttribute("sz")) sz = Xml.intAttr(arPr, "sz", Math.round(size * 100)) / 100f;
            Color c = ctx.theme.colorChild(Xml.child(arPr, "solidFill"), null);
            if (c != null) col = c;
        }
        float em = scaled(sz, st);
        String t = Xml.text(Xml.child(r, "t"));
        StringBuilder ord = new StringBuilder();
        for (int i = 0; i < t.length(); ) {
            int cp = t.codePointAt(i);
            i += Character.charCount(cp);
            if (cp == '-') cp = 0x2212;
            if (cp == '*') cp = 0x2217;
            int k = opKind(cp);
            if (k != 0) {
                if (ord.length() > 0) { items.add(new Text(ord.toString(), font, em, col, ctx.fonts)); ord.setLength(0); }
                Text op = new Text(new String(Character.toChars(cp)), font, em, col, ctx.fonts);
                op.kind = k;
                items.add(op);
            } else ord.appendCodePoint(mathChar(cp, sty));
        }
        if (ord.length() > 0) items.add(new Text(ord.toString(), font, em, col, ctx.fonts));
    }

    /** 演算子の種類 (0 は普通の文字)。 */
    static int opKind(int cp) {
        if ("+−±∓×÷⋅∙∗∘⊕⊗∪∩∧∨".indexOf(cp) >= 0) return 1;
        if ("=≠<>≤≥≈≡∼≅∝→←↔⇒⇔∈∉⊂⊃⊆⊇≪≫".indexOf(cp) >= 0) return 2;
        if ("([{".indexOf(cp) >= 0) return 3;
        if (")]}".indexOf(cp) >= 0) return 4;
        if (",;".indexOf(cp) >= 0) return 5;
        return 0;
    }

    /** 英字・ギリシャ文字を数式の字形にする (sty: p 標準、i 斜体、b 太字、bi 太字の斜体)。 */
    static int mathChar(int cp, String sty) {
        if (sty.equals("p")) return cp;
        boolean upper = cp >= 'A' && cp <= 'Z', lower = cp >= 'a' && cp <= 'z';
        if (upper || lower) {
            int idx = upper ? cp - 'A' : 26 + cp - 'a';
            return switch (sty) {
                case "b" -> 0x1D400 + idx;
                case "bi" -> 0x1D468 + idx;
                default -> cp == 'h' ? 0x210E : 0x1D434 + idx;
            };
        }
        if (cp >= 0x03B1 && cp <= 0x03C9 && !sty.equals("b")) return (sty.equals("bi") ? 0x1D736 : 0x1D6FC) + (cp - 0x03B1);
        if (cp == 0x03B8 && sty.equals("i")) return 0x1D703;
        return cp;
    }
}
