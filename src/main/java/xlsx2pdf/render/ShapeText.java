package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.Element;

/** DrawingML のテキスト本体 (図形の文字列、グラフのタイトル・ラベル) のレイアウトと描画。 */
public final class ShapeText {
    private final RenderContext ctx;
    private final DrawingPainter dp;
    /** 指定がないときの文字サイズ (Excel の図形は 11pt)。 */
    public float defaultSize = 11f;
    /** 指定がないときの文字色。 */
    public Color defaultColor = Color.BLACK;
    /** 指定がないときの太字 (グラフの文字用。null なら標準)。 */
    public Boolean defaultBold;
    /** 段落の既定の配置 (グラフはタイトルを中央揃えにする)。 */
    public String defaultAlgn = "l";

    /** テキストの描画器を作る。 */
    public ShapeText(RenderContext ctx, DrawingPainter dp) {
        this.ctx = ctx;
        this.dp = dp;
    }

    /** 段落。 */
    private static final class Para {
        /** 段落のラン。 */
        List<TextLayout.Run> runs = new ArrayList<>();
        /** 配置 (l, ctr, r, just, dist)。 */
        String algn = "l";
        /** 左余白・字下げ・段落前後の間隔 (pt。負の間隔は行の高さに対する割合)。 */
        double marL, indent, spcBef, spcAft;
        /** 行間 (割合、または pt。pt が正ならそれを使う)。 */
        double lnSpcPct = 1.0, lnSpcPts = -1;
        /** 箇条書きの記号 (なければ null)。 */
        String bullet;
        /** 段落末 (endParaRPr) の文字サイズとフォント (空行の高さに使う)。 */
        float endSize = -1;
        FontManager.Spec endSpec;
    }

    /** リストスタイルとラン自身から積み重ねたランの書式 (null は未指定)。 */
    private static final class RProps {
        Float size;
        Boolean b, i;
        String u, strike;
        Color color;
        String latin, ea;
        Integer baseline;
        Double spc;

        /** 複製する。 */
        RProps copy() {
            RProps p = new RProps();
            p.size = size; p.b = b; p.i = i; p.u = u; p.strike = strike; p.color = color; p.latin = latin; p.ea = ea;
            p.baseline = baseline; p.spc = spc;
            return p;
        }
    }

    /** rPr / defRPr の書式を重ねる。 */
    private void apply(RProps p, Element rPr, Color ph) {
        if (rPr == null) return;
        if (rPr.hasAttribute("sz")) p.size = Xml.intAttr(rPr, "sz", 1100) / 100f;
        if (rPr.hasAttribute("b")) p.b = Xml.boolAttr(rPr, "b", false);
        if (rPr.hasAttribute("i")) p.i = Xml.boolAttr(rPr, "i", false);
        if (rPr.hasAttribute("u")) p.u = rPr.getAttribute("u");
        if (rPr.hasAttribute("strike")) p.strike = rPr.getAttribute("strike");
        if (rPr.hasAttribute("baseline")) p.baseline = Xml.intAttr(rPr, "baseline", 0);
        if (rPr.hasAttribute("spc")) p.spc = Xml.intAttr(rPr, "spc", 0) / 100.0;
        Element sf = Xml.child(rPr, "solidFill");
        if (sf != null) p.color = ctx.theme.colorChild(sf, ph);
        Element gf = Xml.child(rPr, "gradFill");
        if (gf != null) p.color = ctx.theme.colorChild(Xml.child(Xml.child(gf, "gsLst"), "gs"), ph);
        if (Xml.child(rPr, "noFill") != null) p.color = new Color(0, 0, 0, 0);
        Element latin = Xml.child(rPr, "latin"), ea = Xml.child(rPr, "ea");
        if (latin != null && !latin.getAttribute("typeface").isEmpty()) p.latin = latin.getAttribute("typeface");
        if (ea != null && !ea.getAttribute("typeface").isEmpty()) p.ea = ea.getAttribute("typeface");
    }

    /** 書式からランを作る (テーマフォントを解決する)。 */
    private TextLayout.Run mkRun(String text, RProps p, double fontScale) {
        String latin = ctx.theme.fontRef(p.latin == null ? "+mn-lt" : p.latin, false);
        String ea = ctx.theme.fontRef(p.ea == null ? "+mn-ea" : p.ea, true);
        if (ea != null && (ea.isEmpty() || ea.equals(latin))) ea = null;
        boolean bold = p.b != null ? p.b : defaultBold != null && defaultBold;
        boolean italic = p.i != null && p.i;
        FontManager.Spec sp = ctx.fonts.spec(latin, ea, bold, italic);
        float size = (float) ((p.size != null ? p.size : defaultSize) * fontScale);
        TextLayout.Run r = new TextLayout.Run(text, sp, size, p.color != null ? p.color : defaultColor);
        if (p.u != null && !p.u.equals("none")) r.underline = p.u.startsWith("dbl") ? 2 : 1;
        r.strike = p.strike != null && !p.strike.equals("noStrike");
        if (p.baseline != null && p.baseline != 0) r.baselineShift = p.baseline > 0 ? 1 : -1;
        if (p.spc != null) r.spacing = (float) (double) p.spc;
        return r;
    }

    /**
     * txBody を矩形 r (図形の座標) に描く。
     *
     * @param fontRefColor スタイル参照 (fontRef) の文字色 (null 可)
     * @param fontRefFace  スタイル参照のフォント (null 可)
     */
    public void paint(PdfCanvas cv, Element txBody, Rectangle2D r, Color fontRefColor, String fontRefFace) throws IOException {
        paint(cv, txBody, r, fontRefColor, fontRefFace, null);
    }

    /**
     * txBody を矩形 r (図形の座標) に描く。
     *
     * @param extraDefaults 追加の既定の書式 (グラフの txPr の defRPr など。null 可)
     */
    public void paint(PdfCanvas cv, Element txBody, Rectangle2D r, Color fontRefColor, String fontRefFace, Element extraDefaults)
            throws IOException {
        Element bodyPr = Xml.child(txBody, "bodyPr");
        double lIns = Xml.longAttr(bodyPr, "lIns", 91440) / DrawingPainter.EMU;
        double tIns = Xml.longAttr(bodyPr, "tIns", 45720) / DrawingPainter.EMU;
        double rIns = Xml.longAttr(bodyPr, "rIns", 91440) / DrawingPainter.EMU;
        double bIns = Xml.longAttr(bodyPr, "bIns", 45720) / DrawingPainter.EMU;
        String anchor = Xml.attr(bodyPr, "anchor", "t");
        boolean anchorCtr = Xml.boolAttr(bodyPr, "anchorCtr", false);
        boolean wrap = !"none".equals(Xml.attr(bodyPr, "wrap", "square"));
        String vert = Xml.attr(bodyPr, "vert", "horz");
        double bodyRot = Xml.longAttr(bodyPr, "rot", 0) / 60000.0;
        double fontScale = 1, lnReduce = 0;
        Element na = Xml.child(bodyPr, "normAutofit");
        if (na != null) {
            fontScale = Xml.longAttr(na, "fontScale", 100000) / 100000.0;
            lnReduce = Xml.longAttr(na, "lnSpcReduction", 0) / 100000.0;
        }
        if (fontRefColor != null) defaultColor = fontRefColor;

        // リストスタイルのレベル 1 の既定値
        RProps base = new RProps();
        if (fontRefFace != null) base.latin = fontRefFace;
        if (extraDefaults != null) apply(base, extraDefaults, null);
        Element lst = Xml.child(txBody, "lstStyle");
        Element lvl1 = Xml.child(lst, "lvl1pPr");
        if (lvl1 != null) apply(base, Xml.child(lvl1, "defRPr"), null);

        List<Para> paras = new ArrayList<>();
        for (Element p : Xml.children(txBody, "p")) {
            Para para = new Para();
            para.algn = defaultAlgn;
            Element pPr = Xml.child(p, "pPr");
            if (lvl1 != null) applyPPr(para, lvl1);
            if (pPr != null) applyPPr(para, pPr);
            RProps pp = base.copy();
            if (pPr != null) apply(pp, Xml.child(pPr, "defRPr"), null);
            for (Element c : Xml.children(p)) {
                switch (c.getLocalName()) {
                    case "r", "fld" -> {
                        RProps rp = pp.copy();
                        apply(rp, Xml.child(c, "rPr"), null);
                        String t = Xml.text(Xml.child(c, "t"));
                        if (!t.isEmpty()) para.runs.add(mkRun(t.replace('\t', ' '), rp, fontScale));
                    }
                    case "br" -> {
                        RProps rp = pp.copy();
                        apply(rp, Xml.child(c, "rPr"), null);
                        para.runs.add(mkRun("\n", rp, fontScale));
                    }
                    case "endParaRPr" -> {
                        RProps rp = pp.copy();
                        apply(rp, c, null);
                        TextLayout.Run er = mkRun(" ", rp, fontScale);
                        para.endSize = er.size;
                        para.endSpec = er.spec;
                    }
                    default -> { }
                }
            }
            if (para.endSpec == null) {
                TextLayout.Run er = mkRun(" ", pp, fontScale);
                para.endSize = er.size;
                para.endSpec = er.spec;
            }
            if (para.bullet != null && !para.runs.isEmpty()) {
                TextLayout.Run first = para.runs.get(0);
                para.runs.add(0, first.copy(para.bullet + " "));
            }
            paras.add(para);
        }

        boolean vertical = vert.equals("vert") || vert.equals("vert270") || vert.equals("eaVert") || vert.equals("wordArtVert")
                || vert.equals("mongolianVert") || vert.equals("wordArtVertRtl");
        Rectangle2D box = new Rectangle2D.Double(r.getX() + lIns, r.getY() + tIns,
                Math.max(0, r.getWidth() - lIns - rIns), Math.max(0, r.getHeight() - tIns - bIns));
        cv.save();
        // はみ出した文字の切り取り (vertOverflow / horzOverflow が clip・ellipsis)。高さ 0 で clip のテキストボックスは
        // 何も印刷されない
        boolean clipV = !"overflow".equals(Xml.attr(bodyPr, "vertOverflow", "overflow"));
        boolean clipH = !"overflow".equals(Xml.attr(bodyPr, "horzOverflow", "overflow"));
        if (clipV || clipH) {
            double big = 1e5;
            cv.clipRect(clipH ? r.getX() : r.getX() - big, clipV ? r.getY() : r.getY() - big,
                    clipH ? r.getWidth() : r.getWidth() + 2 * big, clipV ? r.getHeight() : r.getHeight() + 2 * big);
        }
        if (bodyRot != 0) {
            AffineTransform at = AffineTransform.getRotateInstance(Math.toRadians(bodyRot), r.getCenterX(), r.getCenterY());
            cv.transform(at);
        }
        if (vertical && !vert.equals("eaVert") && !vert.equals("wordArtVert")) {
            // 枠全体を 90 度回転して横書きでレイアウトする
            double ang = vert.equals("vert270") ? -90 : 90;
            AffineTransform at = new AffineTransform();
            at.translate(box.getCenterX(), box.getCenterY());
            at.rotate(Math.toRadians(ang));
            cv.transform(at);
            Rectangle2D rb = new Rectangle2D.Double(-box.getHeight() / 2, -box.getWidth() / 2, box.getHeight(), box.getWidth());
            layoutHorizontal(cv, paras, rb, anchor, anchorCtr, wrap, lnReduce);
        } else if (vertical) {
            layoutEaVertical(cv, paras, box, anchor);
        } else {
            layoutHorizontal(cv, paras, box, anchor, anchorCtr, wrap, lnReduce);
        }
        cv.restore();
    }

    /** 段落の書式 (配置、余白、行間、段落前後の間隔、箇条書き) を重ねる。 */
    private void applyPPr(Para para, Element pPr) {
        if (pPr.hasAttribute("algn")) para.algn = pPr.getAttribute("algn");
        if (pPr.hasAttribute("marL")) para.marL = Xml.longAttr(pPr, "marL", 0) / DrawingPainter.EMU;
        if (pPr.hasAttribute("indent")) para.indent = Xml.longAttr(pPr, "indent", 0) / DrawingPainter.EMU;
        Element ls = Xml.child(pPr, "lnSpc");
        if (ls != null) {
            Element pct = Xml.child(ls, "spcPct"), pts = Xml.child(ls, "spcPts");
            if (pct != null) para.lnSpcPct = Xml.longAttr(pct, "val", 100000) / 100000.0;
            if (pts != null) para.lnSpcPts = Xml.longAttr(pts, "val", 0) / 100.0;
        }
        para.spcBef = spacing(Xml.child(pPr, "spcBef"), para.spcBef);
        para.spcAft = spacing(Xml.child(pPr, "spcAft"), para.spcAft);
        Element bu = Xml.child(pPr, "buChar");
        if (bu != null) para.bullet = bu.getAttribute("char");
        if (Xml.child(pPr, "buNone") != null) para.bullet = null;
        if (Xml.child(pPr, "buAutoNum") != null) para.bullet = "•";
    }

    /** 段落前後の間隔 (pt。割合指定は負の値で返す)。 */
    private static double spacing(Element e, double dflt) {
        if (e == null) return dflt;
        Element pts = Xml.child(e, "spcPts");
        if (pts != null) return Xml.longAttr(pts, "val", 0) / 100.0;
        Element pct = Xml.child(e, "spcPct");
        if (pct != null) return -Xml.longAttr(pct, "val", 0) / 100000.0; // 負 = 行の高さに対する割合
        return dflt;
    }

    /** DrawingML の 1 行の高さ: フォントのアセント + ディセント (一般的なフォントでサイズの約 1.2 倍)。 */
    static double lineHeight(TextLayout.Line l, double fallbackSize) {
        double h = l.ascent + l.descent;
        if (h <= 0) h = fallbackSize * 1.2;
        return h;
    }

    /**
     * 横書きでレイアウトして描く。
     *
     * @param anchor    縦位置 (t, ctr, b)
     * @param anchorCtr 文字のブロックを横に中央揃えするか (現在は位置に影響しない)
     * @param lnReduce  自動調整による行間の縮小率
     */
    private void layoutHorizontal(PdfCanvas cv, List<Para> paras, Rectangle2D box, String anchor, boolean anchorCtr,
                                  boolean wrap, double lnReduce) throws IOException {
        /** 配置する 1 行 (行送りと、その前の段落間隔)。 */
        record PL(TextLayout.Line line, Para p, double pitch, double before) {}
        List<PL> all = new ArrayList<>();
        double total = 0;
        for (int pi = 0; pi < paras.size(); pi++) {
            Para p = paras.get(pi);
            double avail = box.getWidth() - p.marL;
            List<TextLayout.Line> lines;
            if (p.runs.isEmpty()) {
                lines = new ArrayList<>();
                TextLayout.Line l = new TextLayout.Line();
                l.ascent = p.endSpec.ascent * p.endSize;
                l.descent = p.endSpec.descent * p.endSize;
                l.endsParagraph = true;
                lines.add(l);
            } else {
                lines = TextLayout.layout(p.runs, (float) avail, wrap, true, ctx.fonts);
            }
            for (int li = 0; li < lines.size(); li++) {
                TextLayout.Line l = lines.get(li);
                if (l.runs.isEmpty() && l.ascent == 0) {
                    l.ascent = p.endSpec.ascent * p.endSize;
                    l.descent = p.endSpec.descent * p.endSize;
                }
                double lh = lineHeight(l, p.endSize);
                double pitch = p.lnSpcPts > 0 ? p.lnSpcPts : lh * Math.max(0.1, p.lnSpcPct - lnReduce);
                double before = li == 0 && pi > 0 ? (p.spcBef < 0 ? -p.spcBef * lh : p.spcBef) : 0;
                if (li == 0 && pi > 0) {
                    Para prev = paras.get(pi - 1);
                    before += prev.spcAft < 0 ? -prev.spcAft * lh : prev.spcAft;
                }
                all.add(new PL(l, p, pitch, before));
                total += pitch + before;
            }
        }
        double y = switch (anchor) {
            case "ctr" -> box.getY() + (box.getHeight() - total) / 2;
            case "b" -> box.getY() + box.getHeight() - total;
            default -> box.getY();
        };
        for (PL pl : all) {
            y += pl.before;
            TextLayout.Line l = pl.line;
            double lh = lineHeight(l, pl.p.endSize);
            // ベースライン: 行間の余りは文字の上側に付ける
            double baseline = y + pl.pitch - (lh - l.ascent);
            if (pl.pitch < lh) baseline = y + l.ascent * pl.pitch / lh;
            double x0 = box.getX() + pl.p.marL;
            double aw = box.getWidth() - pl.p.marL;
            double x;
            float extra = 0;
            int mode = 0;
            switch (pl.p.algn) {
                case "ctr" -> x = x0 + (aw - l.width) / 2;
                case "r" -> x = x0 + aw - l.width;
                case "just" -> {
                    x = x0;
                    if (!l.endsParagraph) { extra = (float) (aw - l.width); mode = 1; }
                }
                case "dist" -> {
                    x = x0;
                    extra = (float) Math.max(0, aw - l.width);
                    mode = 2;
                }
                default -> x = x0;
            }
            if (l.runs.isEmpty() || l.width == 0) { y += pl.pitch; continue; }
            if (mode == 1 && l.spaces == 0) {
                // 空白のない (CJK の) 両端揃えは文字の間に配る
                mode = 2;
            }
            TextLayout.drawLine(cv, l, x, baseline, extra, mode);
            y += pl.pitch;
        }
    }

    /**
     * 東アジアの縦書き: 列は右から左、文字は上から下に積む。簡易なレイアウトで、段落ごとに新しい列を始め、
     * 高さを超えたら次の列に送る。
     */
    private void layoutEaVertical(PdfCanvas cv, List<Para> paras, Rectangle2D box, String anchor) throws IOException {
        double x = box.getMaxX();
        /** 配置する 1 文字 (列番号と列内の位置)。 */
        record G(String ch, TextLayout.Run run, double col, double y) {}
        List<G> glyphs = new ArrayList<>();
        int col = 0;
        double y = 0;
        double maxColW = 0;
        for (Para p : paras) {
            y = 0;
            for (TextLayout.Run r : p.runs) {
                for (int i = 0; i < r.text.length(); ) {
                    int cp = r.text.codePointAt(i);
                    i += Character.charCount(cp);
                    if (cp == '\n') { col++; y = 0; continue; }
                    double adv = FontManager.isCjk(cp) ? r.size : r.spec.charWidth(cp, r.size, ctx.fonts);
                    if (y + adv > box.getHeight() + 0.01 && y > 0) { col++; y = 0; }
                    glyphs.add(new G(new String(Character.toChars(cp)), r, col, y));
                    y += adv;
                    maxColW = Math.max(maxColW, r.size * 1.2);
                }
            }
            col++;
        }
        if (maxColW == 0) return;
        for (G g : glyphs) {
            double cx = x - (g.col + 0.5) * maxColW;
            TextLayout.Run r = g.run;
            int cp = g.ch.codePointAt(0);
            if (FontManager.isCjk(cp)) {
                double w = r.spec.charWidth(cp, r.size, ctx.fonts);
                String ch = rotateVerticalForm(g.ch);
                cv.text(ch, cx - w / 2, box.getY() + g.y + r.spec.ascent * r.size * 0.88, r.spec, r.size, r.color, 1, 0);
            } else {
                // 欧文は時計回りに 90 度回転する
                cv.glyph(g.ch, cx - r.size * 0.35, box.getY() + g.y, Math.PI / 2, r.spec, r.size, r.color);
            }
        }
    }

    /** 縦書き用の字形 (長音・括弧・句読点)。 */
    private static String rotateVerticalForm(String ch) {
        return switch (ch) {
            case "ー" -> "丨";
            case "（" -> "︵";
            case "）" -> "︶";
            case "「" -> "﹁";
            case "」" -> "﹂";
            case "、" -> "︑";
            case "。" -> "︒";
            default -> ch;
        };
    }
}
