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
        /** 数式のある段落の組んだ箱 (なければ null) と、数式の段落の配置 (oMathPara の jc。inline は段落の配置)。 */
        MathLayout.Box math;
        String mathJc;
        /** 数式の文字の大きさ (pt)。 */
        float mathSize;
        /** タブ位置 {位置 pt, 種類 (0 左、1 中央、2 右、3 小数点)} (位置の順) と、既定のタブの間隔 (pt)。 */
        List<float[]> tabs = new ArrayList<>();
        double defTab = DEF_TAB;
    }

    /** 既定のタブの間隔 (pt。defTabSz の既定 914400 EMU。Excel も 72pt ごと)。 */
    static final double DEF_TAB = 72;

    /** 段落のタブ位置 (基準は文字の枠の左端。最後のタブ位置より右は既定の間隔の倍数)。 */
    private static TextLayout.TabStops tabStops(Para p) {
        return x -> {
            for (float[] t : p.tabs) if (t[0] > x + 0.01f) return t;
            double d = p.defTab > 0 ? p.defTab : DEF_TAB;
            return new float[] {(float) ((Math.floor(x / d + 1e-6) + 1) * d), 0};
        };
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
        /** 文字の外側の影 (outerShdw の要素。null は未指定)。 */
        Element shadow;

        /** 複製する。 */
        RProps copy() {
            RProps p = new RProps();
            p.size = size; p.b = b; p.i = i; p.u = u; p.strike = strike; p.color = color; p.latin = latin; p.ea = ea;
            p.baseline = baseline; p.spc = spc; p.shadow = shadow;
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
        Element effl = Xml.child(rPr, "effectLst");
        if (effl != null) p.shadow = Xml.child(effl, "outerShdw");
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
        TextLayout.Run r = new TextLayout.Run(text, sp, size, dp != null && dp.bw ? Color.BLACK : p.color != null ? p.color : defaultColor);
        if (p.u != null && !p.u.equals("none")) r.underline = p.u.startsWith("dbl") ? 2 : 1;
        r.strike = p.strike != null && !p.strike.equals("noStrike");
        if (p.baseline != null && p.baseline != 0) r.baselineShift = p.baseline > 0 ? 1 : -1;
        if (p.spc != null) r.spacing = (float) (double) p.spc;
        if (p.shadow != null && !(dp != null && dp.bw)) {
            Color sc = ctx.theme.colorChild(p.shadow, null);
            double dist = Xml.longAttr(p.shadow, "dist", 0) / DrawingPainter.EMU;
            double dir = Math.toRadians(Xml.longAttr(p.shadow, "dir", 0) / 60000.0);
            if (sc != null && dist > 0) { r.shadow = sc; r.shadowDx = (float) (Math.cos(dir) * dist); r.shadowDy = (float) (Math.sin(dir) * dist); }
        }
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
            boolean hasMath = false;
            for (Element c : Xml.children(p)) if (c.getLocalName().equals("m")) hasMath = true;
            if (hasMath) {
                // 数式のある段落: ランと数式を 1 行の箱に並べる
                MathLayout.Group g = new MathLayout.Group();
                double mx = 0;
                for (Element c : Xml.children(p)) {
                    RProps rp = pp.copy();
                    if (c.getLocalName().equals("r")) {
                        apply(rp, Xml.child(c, "rPr"), null);
                        String t = Xml.text(Xml.child(c, "t"));
                        if (t.isEmpty()) continue;
                        TextLayout.Run run = mkRun(t, rp, fontScale);
                        MathLayout.Text tb = new MathLayout.Text(t, run.spec, run.size, run.color, ctx.fonts);
                        g.add(tb, mx, 0);
                        mx += tb.w;
                    } else if (c.getLocalName().equals("m")) {
                        TextLayout.Run run = mkRun(" ", rp, fontScale);
                        float msz = MathLayout.firstSize(c, run.size);
                        MathLayout.Box b = new MathLayout(ctx, msz, run.color).build(c);
                        para.mathSize = Math.max(para.mathSize, msz);
                        g.add(b, mx, 0);
                        mx += b.w;
                        if (para.mathJc == null) para.mathJc = MathLayout.justification(c);
                    }
                }
                para.math = g;
            }
            for (Element c : Xml.children(p)) {
                if (para.math != null && !c.getLocalName().equals("endParaRPr")) continue;
                switch (c.getLocalName()) {
                    case "r", "fld" -> {
                        RProps rp = pp.copy();
                        apply(rp, Xml.child(c, "rPr"), null);
                        String t = Xml.text(Xml.child(c, "t"));
                        if (!t.isEmpty()) para.runs.add(mkRun(t, rp, fontScale));
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

        // タブ位置は横書きの通常の配置だけで使う。変形・東アジアの縦書きではタブを空白にする
        Element warpEl0 = Xml.child(bodyPr, "prstTxWarp");
        boolean eaVert = vert.equals("eaVert") || vert.equals("wordArtVert") || vert.equals("mongolianVert") || vert.equals("wordArtVertRtl");
        if (eaVert || (warpEl0 != null && !"textNoShape".equals(warpEl0.getAttribute("prst")))) {
            for (Para p : paras) for (TextLayout.Run run : p.runs) run.text = run.text.replace('\t', ' ');
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
        Element warpEl = Xml.child(bodyPr, "prstTxWarp");
        String warp = warpEl == null ? null : warpEl.getAttribute("prst");
        if (warp != null && !warp.isEmpty() && !warp.equals("textNoShape") && !vertical && TextWarp.known(warp)) {
            // ワードアートの変形: 文字の輪郭を変形の形に写す。輪郭の線は最初のランの rPr/ln
            List<List<TextLayout.Run>> ls = new ArrayList<>();
            for (Para p : paras) if (!p.runs.isEmpty()) ls.add(p.runs);
            Element rln = null;
            for (Element p : Xml.children(txBody, "p")) {
                for (Element rr : Xml.children(p, "r")) { rln = Xml.child(Xml.child(rr, "rPr"), "ln"); if (rln != null) break; }
                if (rln != null) break;
            }
            DrawingPainter.LineStyle ol = rln == null || Xml.child(rln, "noFill") != null ? null : dp.line(rln, null, null);
            Element shdw = null;
            for (Element p : Xml.children(txBody, "p")) {
                for (Element rr : Xml.children(p, "r")) { shdw = Xml.child(Xml.child(Xml.child(rr, "rPr"), "effectLst"), "outerShdw"); if (shdw != null) break; }
                if (shdw != null) break;
            }
            TextWarp.paint(cv, ctx, ls, r, warp, warpEl, ol, dp, dp != null && dp.bw ? null : shdw); // 変形は図形の矩形全体 (余白を除かない) に合わせる
        } else if (vertical && !vert.equals("eaVert") && !vert.equals("wordArtVert")) {
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
        if (pPr.hasAttribute("defTabSz")) para.defTab = Xml.longAttr(pPr, "defTabSz", 914400) / DrawingPainter.EMU;
        Element tl = Xml.child(pPr, "tabLst");
        if (tl != null) {
            para.tabs = new ArrayList<>();
            for (Element t : Xml.children(tl, "tab")) {
                int ty = switch (Xml.attr(t, "algn", "l")) { case "ctr" -> 1; case "r" -> 2; case "dec" -> 3; default -> 0; };
                para.tabs.add(new float[] {(float) (Xml.longAttr(t, "pos", 0) / DrawingPainter.EMU), ty});
            }
            para.tabs.sort((a, b) -> Float.compare(a[0], b[0]));
        }
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

    private static void textMetrics(TextLayout.Line l) {
        if (l.runs.isEmpty()) return;
        float a = 0, d = 0;
        boolean any = false;
        for (TextLayout.Run r : l.runs) {
            if (r.tabW >= 0) continue;
            boolean cjk = r.text.codePoints().anyMatch(FontManager::isCjk);
            // Excel の図形の行: 和文の書体は 1.3 倍、欧文の書体は hhea + lineGap。欧文の書体の中の和文の文字は代わりの和文の書体の行
            boolean fb = cjk && !r.spec.eaFont;
            // 和文の書体を指定したラン (a:ea) の和文の文字は、その書体の行の高さ
            float ea = cjk && r.spec.eaLineAsc > 0 ? r.spec.eaLineAsc : FontManager.EA_FALLBACK_ASC;
            float ed = cjk && r.spec.eaLineAsc > 0 ? r.spec.eaLineDesc : FontManager.EA_FALLBACK_DESC;
            a = Math.max(a, (fb ? Math.max(ea, r.spec.lineAsc) : r.spec.lineAsc) * r.size);
            d = Math.max(d, (fb ? Math.max(ed, r.spec.lineDesc) : r.spec.lineDesc) * r.size);
            any = true;
        }
        if (any) { l.ascent = a; l.descent = d; }
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
        record PL(TextLayout.Line line, Para p, double pitch, double before, boolean first) {}
        List<PL> all = new ArrayList<>();
        double total = 0;
        for (int pi = 0; pi < paras.size(); pi++) {
            Para p = paras.get(pi);
            double avail = box.getWidth() - p.marL;
            List<TextLayout.Line> lines;
            if (p.math != null) {
                lines = new ArrayList<>();
                TextLayout.Line l = new TextLayout.Line();
                l.ascent = (float) (p.math.asc + MathLayout.PARA_TOP * p.mathSize);
                l.descent = (float) (p.math.dsc + MathLayout.PARA_BOTTOM * p.mathSize);
                l.width = (float) p.math.w;
                l.endsParagraph = true;
                lines.add(l);
            } else if (p.runs.isEmpty()) {
                lines = new ArrayList<>();
                TextLayout.Line l = new TextLayout.Line();
                l.ascent = p.endSpec.lineAsc * p.endSize;
                l.descent = p.endSpec.lineDesc * p.endSize;
                l.endsParagraph = true;
                lines.add(l);
            } else {
                lines = TextLayout.layout(p.runs, (float) avail, wrap, true, ctx.fonts, tabStops(p),
                        (float) (p.marL + p.indent), (float) p.marL);
                for (TextLayout.Line l : lines) textMetrics(l);
            }
            for (int li = 0; li < lines.size(); li++) {
                TextLayout.Line l = lines.get(li);
                if (l.runs.isEmpty() && l.blankBase != null) {
                    // 段落の中の空行 (改行が続く行): 直前の行の先頭のランの書体の行の高さ (和文の 1.3 倍を足さない欧文の行。Excel の実測)
                    l.ascent = l.blankBase.spec.lineAsc * l.blankBase.size;
                    l.descent = l.blankBase.spec.lineDesc * l.blankBase.size;
                } else if (l.runs.isEmpty() && l.ascent == 0) {
                    l.ascent = p.endSpec.lineAsc * p.endSize;
                    l.descent = p.endSpec.lineDesc * p.endSize;
                }
                double lh = lineHeight(l, p.endSize);
                double pitch = p.lnSpcPts > 0 ? p.lnSpcPts : lh * Math.max(0.1, p.lnSpcPct - lnReduce);
                double before = li == 0 && pi > 0 ? (p.spcBef < 0 ? -p.spcBef * lh : p.spcBef) : 0;
                if (li == 0 && pi > 0) {
                    Para prev = paras.get(pi - 1);
                    before += prev.spcAft < 0 ? -prev.spcAft * lh : prev.spcAft;
                }
                all.add(new PL(l, p, pitch, before, li == 0));
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
            // 段落の最初の行は字下げ (indent。負ならぶら下げ) の分だけずらす
            double ind = pl.first ? pl.p.indent : 0;
            double x0 = box.getX() + pl.p.marL + ind;
            double aw = box.getWidth() - pl.p.marL - ind;
            double x;
            float extra = 0;
            int mode = 0;
            switch (pl.p.algn) {
                case "ctr" -> x = x0 + (aw - l.width) / 2;
                case "r" -> x = x0 + aw - l.width;
                case "just" -> {
                    x = x0;
                    // 行が使える幅より長い (折り返さない図形) ときは配らない (負の余りで文字を重ねない)
                    if (!l.endsParagraph) { extra = (float) Math.max(0, aw - l.width); mode = 1; }
                }
                case "dist" -> {
                    x = x0;
                    extra = (float) Math.max(0, aw - l.width);
                    mode = 2;
                }
                default -> x = x0;
            }
            if (pl.p.math != null) {
                // 数式: ディスプレイの数式 (oMathPara) は jc で、行の中の数式は段落の配置で置く
                String jc = pl.p.mathJc == null ? "inline" : pl.p.mathJc;
                double mx = switch (jc) {
                    case "left" -> x0;
                    case "right" -> x0 + aw - l.width;
                    case "center", "centerGroup" -> x0 + (aw - l.width) / 2;
                    default -> x;
                };
                pl.p.math.draw(cv, mx, baseline);
                y += pl.pitch;
                continue;
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
        // 縦書きの anchor は横の位置: t = 右、ctr = 中央、b = 左 (列は右から左に並ぶ)
        double used = col * maxColW;
        if ("ctr".equals(anchor)) x = box.getMaxX() - (box.getWidth() - used) / 2;
        else if ("b".equals(anchor)) x = box.getX() + used;
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
