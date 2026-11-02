package xlsx2pdf.render;

import java.awt.Color;
import java.awt.Shape;
import java.awt.geom.Arc2D;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Line2D;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.w3c.dom.Element;

/** DrawingML のグラフ (棒、折れ線、面、円、ドーナツ、散布図、レーダー) を、キャッシュされたグラフのデータから描く。 */
public final class ChartPainter {
    private final RenderContext ctx;
    private final DrawingPainter dp;
    /** グラフのパーツ。 */
    private final PackagePart part;
    /** グラフ全体の既定の文字書式 (chartSpace の txPr)。 */
    private Element defaultTxPr;

    /**
     * グラフの描画器を作る。
     *
     * @param part グラフのパーツ
     */
    public ChartPainter(RenderContext ctx, DrawingPainter dp, PackagePart part) {
        this.ctx = ctx;
        this.dp = dp;
        this.part = part;
    }

    // ------------------------------------------------------------------ モデル

    /** 系列。 */
    static final class Series {
        /** 系列の番号と順序。 */
        int idx, order;
        /** 系列名。 */
        String name = "";
        /** 項目、値、散布図の x の値。 */
        List<String> cats = new ArrayList<>();
        List<Double> vals = new ArrayList<>();
        List<Double> xs = new ArrayList<>();
        /** 値の表示形式。 */
        String fmt = "General";
        /** 書式・マーカー・データラベル・系列の要素。 */
        Element spPr, marker, dLbls, ser;
        /** 平滑線。 */
        boolean smooth;
        /** 系列の色。 */
        Color color;
    }

    /** グラフの種類ごとのグループ (barChart など)。 */
    static final class Group {
        /** 種類 (bar, line, area, pie, doughnut, scatter, radar)。 */
        String type;
        /** 横棒 (barDir が bar)。 */
        boolean horizontal;
        /** 集合・積み上げ・100% 積み上げ (clustered, stacked, percentStacked, standard)。 */
        String grouping = "clustered";
        /** 要素ごとに色を変える。 */
        boolean varyColors;
        /** 棒の間隔・重なり、ドーナツの穴の大きさ、円の最初の扇の角度。 */
        double gapWidth = 150, overlap = 0, holeSize = 50, firstSliceAng = 0;
        /** 散布図・レーダーのスタイル。 */
        String scatterStyle = "lineMarker";
        String radarStyle = "marker";
        /** 系列。 */
        List<Series> series = new ArrayList<>();
        /** グループのデータラベル。 */
        Element dLbls;
        /** マーカーを表示する。 */
        boolean markers = true;
        /** 使う軸の ID。 */
        List<String> axIds = new ArrayList<>();
    }

    /** 軸。 */
    static final class Axis {
        /** 軸の ID と種類 (valAx, catAx, dateAx)。 */
        String id, kind;
        /** 非表示。 */
        boolean deleted;
        /** 位置 (b, l, t, r)。 */
        String pos = "b";
        /** 軸の要素。 */
        Element el;
        /** 指定の最小・最大・目盛間隔 (null は自動)。 */
        Double min, max, majorUnit;
        /** 軸を反転する。 */
        boolean reverse;
        /** 目盛ラベルの表示形式。 */
        String fmt;
        /** 主・補助の目盛線。 */
        boolean gridMajor, gridMinor;
        Element gridSpPr;
        /** 目盛ラベルの位置。 */
        String tickLblPos = "nextTo";
        /** 交差する位置と、項目の境界での交差。 */
        String crosses = "autoZero";
        String crossBetween = "between";
        /** 交差する軸の ID。 */
        String crossAx;
        /** 計算した目盛りの範囲と間隔。 */
        double lo, hi, unit;
        /** 対数目盛。 */
        boolean logScale;
        /** 軸ラベル。 */
        String title;
        Element titleEl;
    }

    // ------------------------------------------------------------------ 描画

    /** グラフを矩形 r に描く: グラフエリア、タイトル、凡例、プロットエリア。 */
    public void paint(PdfCanvas cv, Rectangle2D r) throws IOException {
        Element root;
        try (InputStream in = part.getInputStream()) {
            root = Xml.parse(in);
        } catch (Exception e) {
            return;
        }
        defaultTxPr = Xml.child(root, "txPr");
        Element chart = Xml.child(root, "chart");
        if (chart == null) return;
        boolean rounded = Xml.boolAttr(Xml.child(root, "roundedCorners"), "val", false);
        if (Xml.child(root, "roundedCorners") == null) rounded = false;

        // グラフエリアの背景
        Element csSpPr = Xml.child(root, "spPr");
        Shape area = rounded ? new java.awt.geom.RoundRectangle2D.Double(r.getX(), r.getY(), r.getWidth(), r.getHeight(), 10, 10) : r;
        DrawingPainter.Paint bg = csSpPr != null ? dp.paint(DrawingPainter.fillElement(csSpPr), null, null) : solid(Color.WHITE);
        if (csSpPr != null && DrawingPainter.fillElement(csSpPr) == null) bg = solid(Color.WHITE);
        dp.fillShape(cv, area, bg, null);
        DrawingPainter.LineStyle border = csSpPr != null ? dp.line(Xml.child(csSpPr, "ln"), null, null) : null;
        if (csSpPr == null || Xml.child(csSpPr, "ln") == null) border = defaultLine(new Color(0xD9D9D9), 0.75);

        cv.save();
        cv.clip(area);
        double pad = 7;
        Rectangle2D inner = new Rectangle2D.Double(r.getX() + pad, r.getY() + pad, r.getWidth() - 2 * pad, r.getHeight() - 2 * pad);

        Element plotArea = Xml.child(chart, "plotArea");
        List<Group> groups = new ArrayList<>();
        List<Axis> axes = new ArrayList<>();
        for (Element e : Xml.children(plotArea)) {
            String n = e.getLocalName();
            if (n.endsWith("Chart")) {
                Group g = parseGroup(e);
                if (g != null) groups.add(g);
            } else if (n.equals("valAx") || n.equals("catAx") || n.equals("dateAx") || n.equals("serAx")) {
                axes.add(parseAxis(e));
            }
        }
        // 既定の色を割り当てる
        int serIndex = 0;
        for (Group g : groups) for (Series s : g.series) s.color = seriesColor(s, Math.max(s.idx, 0), serIndex++);

        // タイトル
        double top = inner.getY();
        Element title = Xml.child(chart, "title");
        boolean autoDeleted = Xml.boolAttr(Xml.child(chart, "autoTitleDeleted"), "val", false);
        String autoTitle = null;
        if (title == null && !autoDeleted && groups.size() == 1 && groups.get(0).series.size() == 1
                && (groups.get(0).type.equals("pie") || groups.get(0).type.equals("doughnut"))) {
            autoTitle = groups.get(0).series.get(0).name;
        }
        if (title != null || autoTitle != null) {
            String fallback = title != null ? null : autoTitle;
            if (title != null && Xml.child(title, "tx") == null && groups.size() > 0 && groups.get(0).series.size() == 1)
                fallback = groups.get(0).series.get(0).name;
            if (title != null && Xml.child(title, "tx") == null && fallback == null) fallback = "グラフ タイトル";
            double h = paintTitle(cv, title, fallback, inner, 14, true);
            boolean overlay = title != null && Xml.boolAttr(Xml.child(title, "overlay"), "val", false);
            if (!overlay) top += h + 4;
        }

        // 凡例
        Element legend = Xml.child(chart, "legend");
        Rectangle2D plotBox = new Rectangle2D.Double(inner.getX(), top, inner.getWidth(), inner.getMaxY() - top);
        if (legend != null) plotBox = paintLegend(cv, legend, groups, plotBox, inner);

        // プロットエリアの手動レイアウト
        Element layout = Xml.path(plotArea, "layout", "manualLayout");
        boolean innerLayout = false;
        if (layout != null && Xml.child(layout, "x") != null) {
            double x = Xml.doubleAttr(Xml.child(layout, "x"), "val", 0), y = Xml.doubleAttr(Xml.child(layout, "y"), "val", 0);
            double w = Xml.doubleAttr(Xml.child(layout, "w"), "val", 1), h = Xml.doubleAttr(Xml.child(layout, "h"), "val", 1);
            plotBox = new Rectangle2D.Double(r.getX() + x * r.getWidth(), r.getY() + y * r.getHeight(), w * r.getWidth(), h * r.getHeight());
            innerLayout = "inner".equals(Xml.attr(Xml.child(layout, "layoutTarget"), "val", "outer"));
        }

        if (groups.isEmpty()) {
            cv.restore();
            if (border != null) dp.strokeShape(cv, area, border);
            return;
        }
        boolean radial = groups.get(0).type.equals("pie") || groups.get(0).type.equals("doughnut");
        if (radial) {
            paintPie(cv, groups.get(0), plotBox);
        } else if (groups.get(0).type.equals("radar")) {
            paintRadar(cv, groups.get(0), axes, plotBox);
        } else {
            paintCartesian(cv, groups, axes, plotBox, innerLayout, Xml.child(plotArea, "spPr"));
        }
        cv.restore();
        if (border != null) dp.strokeShape(cv, area, border);
    }

    /** 単色の塗りつぶし。 */
    private static DrawingPainter.Paint solid(Color c) {
        DrawingPainter.Paint p = new DrawingPainter.Paint();
        p.color = c;
        return p;
    }

    /** 指定の色・太さの実線。 */
    private static DrawingPainter.LineStyle defaultLine(Color c, double w) {
        DrawingPainter.LineStyle l = new DrawingPainter.LineStyle();
        l.color = c;
        l.width = w;
        return l;
    }

    // ------------------------------------------------------------------ 解析

    /** グラフのグループ要素を解析する。 */
    private Group parseGroup(Element e) {
        Group g = new Group();
        String n = e.getLocalName();
        switch (n) {
            case "barChart", "bar3DChart" -> {
                g.type = "bar";
                g.horizontal = "bar".equals(Xml.attr(Xml.child(e, "barDir"), "val", "col"));
                g.gapWidth = Xml.doubleAttr(Xml.child(e, "gapWidth"), "val", 150);
                g.overlap = Xml.doubleAttr(Xml.child(e, "overlap"), "val", 0);
            }
            case "lineChart", "line3DChart", "stockChart" -> g.type = "line";
            case "areaChart", "area3DChart" -> g.type = "area";
            case "pieChart", "pie3DChart", "ofPieChart" -> g.type = "pie";
            case "doughnutChart" -> {
                g.type = "doughnut";
                g.holeSize = Xml.doubleAttr(Xml.child(e, "holeSize"), "val", 50);
            }
            case "scatterChart", "bubbleChart" -> {
                g.type = "scatter";
                g.scatterStyle = Xml.attr(Xml.child(e, "scatterStyle"), "val", "lineMarker");
            }
            case "radarChart" -> {
                g.type = "radar";
                g.radarStyle = Xml.attr(Xml.child(e, "radarStyle"), "val", "marker");
            }
            default -> { return null; }
        }
        g.grouping = Xml.attr(Xml.child(e, "grouping"), "val", g.type.equals("bar") ? "clustered" : "standard");
        if (g.type.equals("bar") && !g.grouping.equals("clustered") && Xml.child(e, "overlap") == null) g.overlap = 100;
        g.varyColors = Xml.boolAttr(Xml.child(e, "varyColors"), "val", false);
        if (Xml.child(e, "varyColors") == null) g.varyColors = g.type.equals("pie") || g.type.equals("doughnut");
        g.firstSliceAng = Xml.doubleAttr(Xml.child(e, "firstSliceAng"), "val", 0);
        g.dLbls = Xml.child(e, "dLbls");
        Element mk = Xml.child(e, "marker");
        g.markers = mk == null || Xml.boolAttr(mk, "val", true);
        for (Element ax : Xml.children(e, "axId")) g.axIds.add(ax.getAttribute("val"));
        for (Element s : Xml.children(e, "ser")) g.series.add(parseSeries(s));
        g.series.sort((a, b) -> Integer.compare(a.order, b.order));
        return g;
    }

    /** 系列を解析する。 */
    private Series parseSeries(Element s) {
        Series se = new Series();
        se.ser = s;
        se.idx = Xml.intAttr(Xml.child(s, "idx"), "val", 0);
        se.order = Xml.intAttr(Xml.child(s, "order"), "val", se.idx);
        Element tx = Xml.child(s, "tx");
        if (tx != null) {
            List<String> pts = strings(Xml.firstChild(tx));
            if (!pts.isEmpty()) se.name = pts.get(0);
            Element v = Xml.child(tx, "v");
            if (v != null) se.name = Xml.text(v);
        }
        if (se.name.isEmpty()) se.name = "系列" + (se.idx + 1);
        se.spPr = Xml.child(s, "spPr");
        se.marker = Xml.child(s, "marker");
        se.dLbls = Xml.child(s, "dLbls");
        se.smooth = Xml.boolAttr(Xml.child(s, "smooth"), "val", false);
        Element cat = Xml.child(s, "cat");
        if (cat == null) cat = Xml.child(s, "xVal");
        Element val = Xml.child(s, "val");
        if (val == null) val = Xml.child(s, "yVal");
        if (cat != null) {
            Element ref = Xml.firstChild(cat);
            se.cats = strings(ref);
            if (Xml.child(s, "xVal") != null) {
                List<Double> xs = numbers(ref);
                if (!xs.isEmpty()) se.xs = xs;
            }
        }
        if (val != null) {
            Element ref = Xml.firstChild(val);
            se.vals = numbers(ref);
            Element cache = Xml.child(ref, "numCache");
            if (cache != null) se.fmt = Xml.text(Xml.child(cache, "formatCode"));
        }
        return se;
    }

    /** strCache / numCache / multiLvlStrCache (最後の階層)、またはリテラルの値の文字列。 */
    private List<String> strings(Element ref) {
        List<String> out = new ArrayList<>();
        if (ref == null) return out;
        Element cache = null;
        for (String n : new String[] {"strCache", "numCache", "multiLvlStrCache", "strLit", "numLit"}) {
            cache = Xml.child(ref, n);
            if (cache != null) break;
        }
        if (cache == null && (ref.getLocalName().equals("strLit") || ref.getLocalName().equals("numLit"))) cache = ref;
        if (cache == null) return out;
        Element lvlParent = cache;
        if (cache.getLocalName().equals("multiLvlStrCache")) {
            List<Element> lvls = Xml.children(cache, "lvl");
            if (lvls.isEmpty()) return out;
            lvlParent = lvls.get(0);
        }
        int count = Xml.intAttr(Xml.child(cache, "ptCount"), "val", 0);
        String fmt = Xml.text(Xml.child(cache, "formatCode"));
        String[] arr = new String[count];
        for (Element pt : Xml.children(lvlParent, "pt")) {
            int i = Xml.intAttr(pt, "idx", 0);
            String v = Xml.text(Xml.child(pt, "v"));
            if (cache.getLocalName().equals("numCache") && !fmt.isEmpty()) {
                try {
                    v = ctx.formatter.formatNumber(Double.parseDouble(v), fmt);
                } catch (NumberFormatException ignored) {
                    // 文字列のまま
                }
            }
            if (i >= 0 && i < count) arr[i] = v;
            else if (i >= count) {
                out.add(v);
            }
        }
        for (String s : arr) out.add(s == null ? "" : s);
        return out;
    }

    /** 数値の参照またはリテラルの値 (数値でなければ null)。 */
    private List<Double> numbers(Element ref) {
        List<Double> out = new ArrayList<>();
        if (ref == null) return out;
        Element cache = Xml.child(ref, "numCache");
        if (cache == null && ref.getLocalName().equals("numLit")) cache = ref;
        if (cache == null) return out;
        int count = Xml.intAttr(Xml.child(cache, "ptCount"), "val", 0);
        Double[] arr = new Double[count];
        for (Element pt : Xml.children(cache, "pt")) {
            int i = Xml.intAttr(pt, "idx", 0);
            try {
                if (i >= 0 && i < count) arr[i] = Double.parseDouble(Xml.text(Xml.child(pt, "v")));
            } catch (NumberFormatException ignored) {
                // 空
            }
        }
        for (Double d : arr) out.add(d);
        return out;
    }

    /** 軸を解析する。 */
    private Axis parseAxis(Element e) {
        Axis a = new Axis();
        a.el = e;
        a.kind = e.getLocalName();
        a.id = Xml.attr(Xml.child(e, "axId"), "val", "");
        a.deleted = Xml.boolAttr(Xml.child(e, "delete"), "val", false);
        a.pos = Xml.attr(Xml.child(e, "axPos"), "val", "b");
        Element sc = Xml.child(e, "scaling");
        if (sc != null) {
            a.reverse = "maxMin".equals(Xml.attr(Xml.child(sc, "orientation"), "val", "minMax"));
            if (Xml.child(sc, "min") != null) a.min = Xml.doubleAttr(Xml.child(sc, "min"), "val", 0);
            if (Xml.child(sc, "max") != null) a.max = Xml.doubleAttr(Xml.child(sc, "max"), "val", 0);
            a.logScale = Xml.child(sc, "logBase") != null;
        }
        if (Xml.child(e, "majorUnit") != null) a.majorUnit = Xml.doubleAttr(Xml.child(e, "majorUnit"), "val", 1);
        Element nf = Xml.child(e, "numFmt");
        if (nf != null && !Xml.boolAttr(nf, "sourceLinked", false)) a.fmt = nf.getAttribute("formatCode");
        Element g = Xml.child(e, "majorGridlines");
        a.gridMajor = g != null;
        a.gridSpPr = g == null ? null : Xml.child(g, "spPr");
        a.gridMinor = Xml.child(e, "minorGridlines") != null;
        a.tickLblPos = Xml.attr(Xml.child(e, "tickLblPos"), "val", "nextTo");
        a.crosses = Xml.attr(Xml.child(e, "crosses"), "val", "autoZero");
        a.crossBetween = Xml.attr(Xml.child(e, "crossBetween"), "val", "between");
        a.crossAx = Xml.attr(Xml.child(e, "crossAx"), "val", "");
        a.titleEl = Xml.child(e, "title");
        return a;
    }

    // ------------------------------------------------------------------ 色と文字

    /** 系列の色 (明示の指定、なければ既定の色)。 */
    private Color seriesColor(Series s, int idx, int pos) {
        Element fill = s.spPr == null ? null : DrawingPainter.fillElement(s.spPr);
        if (fill != null) {
            DrawingPainter.Paint p = dp.paint(fill, null, null);
            if (p != null && p.color != null) return p.color;
        }
        if (s.spPr != null) {
            Element ln = Xml.child(s.spPr, "ln");
            Color lc = ln == null ? null : ctx.theme.colorChild(Xml.child(ln, "solidFill"), null);
            if (lc != null && fill == null) return lc;
        }
        return autoColor(idx);
    }

    /** Excel の既定の系列の色: accent1〜6、その後は暗い・明るい変化形。 */
    Color autoColor(int idx) {
        String[] acc = {"accent1", "accent2", "accent3", "accent4", "accent5", "accent6"};
        Color base = ctx.theme.scheme.get(acc[idx % 6]);
        if (base == null) base = Color.GRAY;
        int cycle = idx / 6;
        if (cycle == 0) return base;
        double[] mods = {0.6, 0.8 /* lighter */, 0.5};
        double[] hls = Theme.rgbToHls(base);
        if (cycle % 2 == 1) hls[1] = hls[1] * mods[0];
        else hls[1] = hls[1] * 0.8 + 0.2;
        return Theme.hlsToRgb(hls, 255);
    }

    /** データ要素の色 (要素ごとの指定、要素ごとに色を変える設定、系列の色の順)。 */
    private Color pointColor(Group g, Series s, int i) {
        // データ要素ごとの明示の指定
        for (Element dpt : Xml.children(s.ser, "dPt")) {
            if (Xml.intAttr(Xml.child(dpt, "idx"), "val", -1) == i) {
                Element spPr = Xml.child(dpt, "spPr");
                Element f = spPr == null ? null : DrawingPainter.fillElement(spPr);
                if (f != null) {
                    DrawingPainter.Paint p = dp.paint(f, null, null);
                    if (p != null && p.color != null) return p.color;
                }
            }
        }
        if (g.varyColors && g.series.size() == 1) return autoColor(i);
        return s.color;
    }

    /** txPr の連なりから決まる文字の書式。 */
    private final class TStyle {
        /** サイズ (pt)。 */
        float size = 10;
        /** 太字。 */
        boolean bold;
        /** 色 (既定は Excel のグラフの文字の灰色)。 */
        Color color = new Color(0x595959);
        /** 欧文・東アジアのフォント (テーマ参照可)。 */
        String latin, ea;

        /** txPr の既定の書式 (defRPr) を重ねる。 */
        TStyle apply(Element txPr) {
            if (txPr == null) return this;
            Element p = Xml.child(txPr, "p");
            Element ppr = Xml.child(p, "pPr");
            Element def = Xml.child(ppr, "defRPr");
            applyRPr(def);
            return this;
        }

        /** rPr / defRPr を重ねる。 */
        void applyRPr(Element rpr) {
            if (rpr == null) return;
            if (rpr.hasAttribute("sz")) size = Xml.intAttr(rpr, "sz", 1000) / 100f;
            if (rpr.hasAttribute("b")) bold = Xml.boolAttr(rpr, "b", false);
            Color c = ctx.theme.colorChild(Xml.child(rpr, "solidFill"), null);
            if (c != null) color = c;
            Element l = Xml.child(rpr, "latin"), e = Xml.child(rpr, "ea");
            if (l != null) latin = l.getAttribute("typeface");
            if (e != null) ea = e.getAttribute("typeface");
        }

        /** フォント。 */
        FontManager.Spec spec() {
            String lt = ctx.theme.fontRef(latin == null ? "+mn-lt" : latin, false);
            String et = ctx.theme.fontRef(ea == null ? "+mn-ea" : ea, true);
            return ctx.fonts.spec(lt, et, bold, false);
        }
    }

    /** グラフの既定の書式に txPr を重ねた書式。 */
    private TStyle style(Element txPr) {
        TStyle t = new TStyle();
        t.apply(defaultTxPr);
        t.apply(txPr);
        return t;
    }

    /** 文字列の幅 (pt)。 */
    private double textWidth(String s, TStyle t) {
        return t.spec().width(s, t.size, ctx.fonts);
    }

    /** 文字列を基準位置に揃えて描く。 */
    private void text(PdfCanvas cv, String s, double x, double y, TStyle t, int halign, int valign) throws IOException {
        // halign: -1 左、0 中央、1 右。valign: -1 上、0 中央、1 下 (y が基準)
        FontManager.Spec sp = t.spec();
        double w = sp.width(s, t.size, ctx.fonts);
        double asc = sp.ascent * t.size, desc = sp.descent * t.size;
        double xx = halign < 0 ? x : halign == 0 ? x - w / 2 : x - w;
        double base = valign < 0 ? y + asc : valign == 0 ? y + (asc - desc) / 2 : y - desc;
        cv.text(s, xx, base, sp, t.size, t.color, 1, 0);
    }

    /** 1 行の高さ (pt)。 */
    private double lineH(TStyle t) {
        FontManager.Spec sp = t.spec();
        return (sp.ascent + sp.descent) * t.size;
    }

    /**
     * タイトルを矩形の上中央に描く。
     *
     * @return 使った高さ (pt)
     */
    private double paintTitle(PdfCanvas cv, Element title, String fallback, Rectangle2D box, float dfltSize, boolean dfltBold)
            throws IOException {
        TStyle t = style(null);
        t.size = dfltSize;
        t.bold = false;
        t.color = new Color(0x595959);
        if (title != null) t.apply(Xml.child(title, "txPr"));
        Element rich = title == null ? null : Xml.path(title, "tx", "rich");
        if (rich != null) {
            // 書式付きのタイトルはランの書式を保つため ShapeText で描く
            ShapeText st = new ShapeText(ctx, dp);
            st.defaultSize = t.size;
            st.defaultColor = t.color;
            st.defaultBold = t.bold;
            st.defaultAlgn = "ctr";
            // 高さの概算: 段落数 × 行の高さ
            int lines = Xml.children(rich, "p").size();
            double maxSize = t.size;
            for (Element p : Xml.children(rich, "p")) {
                for (Element rr : Xml.children(p, "r")) {
                    Element rpr = Xml.child(rr, "rPr");
                    if (rpr != null && rpr.hasAttribute("sz")) maxSize = Math.max(maxSize, Xml.intAttr(rpr, "sz", 1400) / 100.0);
                }
            }
            double h = lines * maxSize * 1.25;
            Rectangle2D tr = manualTitleRect(title, box, h);
            st.paint(cv, rich, tr, t.color, null, null);
            return h;
        }
        String s = fallback;
        if (title != null) {
            Element sref = Xml.path(title, "tx", "strRef");
            if (sref != null) {
                List<String> v = strings(sref);
                if (!v.isEmpty()) s = String.join(" ", v);
            }
        }
        if (s == null || s.isEmpty()) return 0;
        double h = lineH(t);
        text(cv, s, box.getCenterX(), box.getY(), t, 0, -1);
        return h;
    }

    /** タイトルの手動レイアウトの矩形。 */
    private Rectangle2D manualTitleRect(Element title, Rectangle2D box, double h) {
        return new Rectangle2D.Double(box.getX(), box.getY(), box.getWidth(), h + 2);
    }

    // ------------------------------------------------------------------ 凡例

    /**
     * 凡例を描く。
     *
     * @return 凡例を除いたプロットエリアの矩形
     */
    private Rectangle2D paintLegend(PdfCanvas cv, Element legend, List<Group> groups, Rectangle2D plot, Rectangle2D inner)
            throws IOException {
        String pos = Xml.attr(Xml.child(legend, "legendPos"), "val", "r");
        boolean overlay = Xml.boolAttr(Xml.child(legend, "overlay"), "val", false);
        TStyle t = style(Xml.child(legend, "txPr"));
        List<String> names = new ArrayList<>();
        List<Color> colors = new ArrayList<>();
        List<Boolean> lineSym = new ArrayList<>();
        java.util.Set<Integer> deleted = new java.util.HashSet<>();
        for (Element le : Xml.children(legend, "legendEntry")) {
            if (Xml.boolAttr(Xml.child(le, "delete"), "val", false)) deleted.add(Xml.intAttr(Xml.child(le, "idx"), "val", -1));
        }
        Group g0 = groups.get(0);
        if ((g0.type.equals("pie") || g0.type.equals("doughnut")) || (g0.varyColors && g0.series.size() == 1 && !g0.series.isEmpty())) {
            Series s = g0.series.isEmpty() ? null : g0.series.get(0);
            if (s != null) for (int i = 0; i < s.cats.size() || i < s.vals.size(); i++) {
                if (deleted.contains(i)) continue;
                names.add(i < s.cats.size() ? s.cats.get(i) : String.valueOf(i + 1));
                colors.add(pointColor(g0, s, i));
                lineSym.add(false);
            }
        } else {
            int k = 0;
            for (Group g : groups) {
                for (Series s : g.series) {
                    if (!deleted.contains(s.idx)) {
                        names.add(s.name);
                        colors.add(s.color);
                        lineSym.add(g.type.equals("line") || (g.type.equals("scatter") && !g.scatterStyle.equals("marker")) || g.type.equals("radar"));
                    }
                    k++;
                }
            }
        }
        if (names.isEmpty()) return plot;
        double lh = lineH(t) * 1.05;
        double sym = t.size * 0.7;
        double symW = sym;
        double maxW = 0;
        for (String n : names) maxW = Math.max(maxW, textWidth(n, t));
        boolean vertical = pos.equals("r") || pos.equals("l") || pos.equals("tr");
        Element ml = Xml.path(legend, "layout", "manualLayout");
        Rectangle2D box;
        if (vertical) {
            double w = symW + 4 + maxW + 6;
            double h = names.size() * lh + 4;
            double x = pos.equals("l") ? inner.getX() : inner.getMaxX() - w;
            double y = plot.getY() + (plot.getHeight() - h) / 2;
            box = new Rectangle2D.Double(x, y, w, h);
            if (!overlay) {
                if (pos.equals("l")) plot = new Rectangle2D.Double(plot.getX() + w + 6, plot.getY(), plot.getWidth() - w - 6, plot.getHeight());
                else plot = new Rectangle2D.Double(plot.getX(), plot.getY(), plot.getWidth() - w - 6, plot.getHeight());
            }
        } else {
            double total = 0;
            for (String n : names) total += symW + 4 + textWidth(n, t) + 10;
            double h = lh + 4;
            double y = pos.equals("t") ? plot.getY() : inner.getMaxY() - h;
            box = new Rectangle2D.Double(inner.getCenterX() - total / 2, y, total, h);
            if (!overlay) {
                if (pos.equals("t")) plot = new Rectangle2D.Double(plot.getX(), plot.getY() + h + 4, plot.getWidth(), plot.getHeight() - h - 4);
                else plot = new Rectangle2D.Double(plot.getX(), plot.getY(), plot.getWidth(), plot.getHeight() - h - 4);
            }
        }
        if (ml != null && Xml.child(ml, "x") != null) {
            Rectangle2D cr = cv == null ? inner : inner;
            double x = Xml.doubleAttr(Xml.child(ml, "x"), "val", 0), y = Xml.doubleAttr(Xml.child(ml, "y"), "val", 0);
            Rectangle2D outer = new Rectangle2D.Double(inner.getX() - 7, inner.getY() - 7, inner.getWidth() + 14, inner.getHeight() + 14);
            double w = Xml.child(ml, "w") != null ? Xml.doubleAttr(Xml.child(ml, "w"), "val", 0) * outer.getWidth() : box.getWidth();
            double h = Xml.child(ml, "h") != null ? Xml.doubleAttr(Xml.child(ml, "h"), "val", 0) * outer.getHeight() : box.getHeight();
            box = new Rectangle2D.Double(outer.getX() + x * outer.getWidth(), outer.getY() + y * outer.getHeight(), w, h);
            if (cr == null) box = inner;
        }
        Element spPr = Xml.child(legend, "spPr");
        if (spPr != null) {
            dp.fillShape(cv, box, dp.paint(DrawingPainter.fillElement(spPr), null, null), null);
            dp.strokeShape(cv, box, dp.line(Xml.child(spPr, "ln"), null, null));
        }
        double x = box.getX() + 3, y = box.getY() + 2;
        for (int i = 0; i < names.size(); i++) {
            if (vertical) {
                drawSymbol(cv, x, y + lh / 2, symW, sym, colors.get(i), lineSym.get(i));
                text(cv, names.get(i), x + symW + 4, y + lh / 2, t, -1, 0);
                y += lh;
            } else {
                drawSymbol(cv, x, y + lh / 2, symW, sym, colors.get(i), lineSym.get(i));
                text(cv, names.get(i), x + symW + 4, y + lh / 2, t, -1, 0);
                x += symW + 4 + textWidth(names.get(i), t) + 10;
            }
        }
        return plot;
    }

    /** 凡例の記号 (線またはマーカー、四角) を描く。 */
    private void drawSymbol(PdfCanvas cv, double x, double cy, double w, double h, Color c, boolean line) throws IOException {
        if (line) {
            cv.line(x, cy, x + w * 1.6, cy, new PdfCanvas.Stroke(2f), c);
        } else {
            cv.fillRect(x, cy - h / 2, w, h, c);
        }
    }

    // ------------------------------------------------------------------ 軸

    /** 目盛間隔: 範囲 ÷ 目標の目盛数 以上で最小の 1, 2, 5, 10 × 10^n。 */
    static double niceUnit(double range, int targetTicks) {
        if (range <= 0) return 1;
        double raw = range / targetTicks;
        double mag = Math.pow(10, Math.floor(Math.log10(raw)));
        double n = raw / mag;
        double u = n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10;
        return u * mag;
    }

    /** Excel に似た値の軸の自動目盛り。 */
    static void autoScale(Axis a, double dmin, double dmax, double lengthPt) {
        if (Double.isInfinite(dmin)) { dmin = 0; dmax = 1; }
        double lo = a.min != null ? a.min : Math.min(0, dmin);
        double hi = a.max != null ? a.max : Math.max(0, dmax);
        if (a.min == null && dmin > 0 && (dmax - dmin) / dmax < 1.0 / 6) lo = dmin - (dmax - dmin) / 2; // Excel は最小値を引き上げる
        if (hi == lo) hi = lo + 1;
        int target = Math.max(2, (int) Math.min(10, lengthPt / 28));
        double unit = a.majorUnit != null ? a.majorUnit : niceUnit(hi - lo, target);
        if (a.min == null) lo = Math.floor(lo / unit) * unit;
        if (a.max == null) {
            double h2 = Math.ceil(hi / unit) * unit;
            if (dmax > 0 && h2 - dmax < 0.05 * (h2 - lo)) h2 += unit; // 上端の余裕
            hi = h2;
        }
        if (a.min == null && dmin < 0) {
            double l2 = Math.floor(dmin / unit) * unit;
            if (dmin < 0 && dmin - l2 < 0.05 * (hi - l2)) l2 -= unit;
            lo = l2;
        }
        a.lo = lo;
        a.hi = hi;
        a.unit = unit;
    }

    /** 軸の目盛ラベルの文字列。 */
    private String axisLabel(Axis a, double v, String srcFmt) {
        String fmt = a.fmt != null ? a.fmt : srcFmt;
        if (Math.abs(v) < 1e-12) v = 0;
        return ctx.formatter.formatNumber(v, fmt == null ? "General" : fmt);
    }

    /** 直交座標のグラフ (棒、折れ線、面、散布図) を描く: 目盛り、目盛線、系列、軸、ラベル。 */
    private void paintCartesian(PdfCanvas cv, List<Group> groups, List<Axis> axes, Rectangle2D box, boolean innerLayout,
                                Element plotSpPr) throws IOException {
        Group g = groups.get(0);
        boolean scatter = g.type.equals("scatter");
        boolean horiz = g.type.equals("bar") && g.horizontal;
        Axis catAx = null, valAx = null, xValAx = null;
        for (Axis a : axes) {
            if (!g.axIds.isEmpty() && !g.axIds.contains(a.id)) continue;
            if (a.kind.equals("catAx") || a.kind.equals("dateAx")) catAx = a;
            else if (a.kind.equals("valAx")) {
                if (scatter && (a.pos.equals("b") || a.pos.equals("t"))) xValAx = a;
                else if (valAx == null) valAx = a;
                else xValAx = a;
            }
        }
        if (valAx == null) { valAx = new Axis(); valAx.kind = "valAx"; valAx.pos = horiz ? "b" : "l"; }
        if (catAx == null && !scatter) { catAx = new Axis(); catAx.kind = "catAx"; catAx.pos = horiz ? "l" : "b"; }
        if (scatter && xValAx == null) { xValAx = new Axis(); xValAx.kind = "valAx"; xValAx.pos = "b"; }

        // データの範囲
        int nCats = 0;
        for (Group gg : groups) for (Series s : gg.series) nCats = Math.max(nCats, Math.max(s.vals.size(), s.cats.size()));
        double dmin = Double.POSITIVE_INFINITY, dmax = Double.NEGATIVE_INFINITY;
        for (Group gg : groups) {
            boolean stacked = gg.grouping.equals("stacked") || gg.grouping.equals("percentStacked");
            if (gg.grouping.equals("percentStacked")) { dmin = Math.min(dmin, 0); dmax = Math.max(dmax, 1); continue; }
            for (int i = 0; i < nCats; i++) {
                double pos = 0, neg = 0;
                for (Series s : gg.series) {
                    Double v = i < s.vals.size() ? s.vals.get(i) : null;
                    if (v == null) continue;
                    if (stacked) { if (v >= 0) pos += v; else neg += v; }
                    else { dmin = Math.min(dmin, v); dmax = Math.max(dmax, v); }
                }
                if (stacked) { dmin = Math.min(dmin, neg); dmax = Math.max(dmax, pos); }
            }
        }
        String srcFmt = g.series.isEmpty() ? "General" : g.series.get(0).fmt;
        if (g.grouping.equals("percentStacked")) srcFmt = "0%";

        TStyle vt = style(Xml.child(valAx.el, "txPr"));
        TStyle ct = catAx != null ? style(Xml.child(catAx.el, "txPr")) : vt;
        if (scatter) ct = style(Xml.child(xValAx.el, "txPr"));

        // ラベルの幅を測るための仮の目盛り
        double valLen = horiz ? box.getWidth() : box.getHeight();
        autoScale(valAx, dmin, dmax, valLen);
        double xmin = Double.POSITIVE_INFINITY, xmax = Double.NEGATIVE_INFINITY;
        if (scatter) {
            for (Series s : g.series) {
                for (int i = 0; i < s.vals.size(); i++) {
                    double x = i < s.xs.size() && s.xs.get(i) != null ? s.xs.get(i) : i + 1;
                    xmin = Math.min(xmin, x);
                    xmax = Math.max(xmax, x);
                }
            }
            autoScale(xValAx, xmin, xmax, box.getWidth());
        }

        // ラベルの場所を空ける
        Rectangle2D plot = box;
        if (!innerLayout) {
            double left = 0, bottom = 0;
            boolean valVisible = !valAx.deleted && !"none".equals(valAx.tickLblPos);
            boolean catVisible = scatter ? !xValAx.deleted : catAx != null && !catAx.deleted && !"none".equals(catAx.tickLblPos);
            if (!horiz) {
                if (valVisible) {
                    double mw = 0;
                    for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v += valAx.unit) mw = Math.max(mw, textWidth(axisLabel(valAx, v, srcFmt), vt));
                    left = mw + 6;
                }
                if (catVisible) bottom = lineH(ct) + 5;
                if (catVisible && !scatter) bottom = catLabelHeight(g, catAx, ct, (box.getWidth() - left) / Math.max(1, nCats)) + 5;
            } else {
                if (catVisible) {
                    double mw = 0;
                    Series s0 = g.series.isEmpty() ? null : g.series.get(0);
                    if (s0 != null) for (String c : s0.cats) mw = Math.max(mw, textWidth(c, ct));
                    left = mw + 6;
                }
                if (valVisible) bottom = lineH(vt) + 5;
            }
            if (valAx.titleEl != null) left += lineH(vt) + 4;
            plot = new Rectangle2D.Double(box.getX() + left, box.getY() + 4, Math.max(10, box.getWidth() - left - 4),
                    Math.max(10, box.getHeight() - bottom - 4));
            autoScale(valAx, dmin, dmax, horiz ? plot.getWidth() : plot.getHeight());
        }

        // プロットエリアの塗りつぶし
        if (plotSpPr != null) {
            dp.fillShape(cv, plot, dp.paint(DrawingPainter.fillElement(plotSpPr), null, null), null);
        }
        // 目盛線
        DrawingPainter.LineStyle grid = valAx.gridMajor ? dp.line(valAx.gridSpPr == null ? null : Xml.child(valAx.gridSpPr, "ln"), null, null) : null;
        if (valAx.gridMajor && grid == null && (valAx.gridSpPr == null || Xml.child(valAx.gridSpPr, "ln") == null)) grid = defaultLine(new Color(0xD9D9D9), 0.75);
        for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v += valAx.unit) {
            double p = valPos(valAx, v, plot, horiz);
            if (grid != null) {
                Line2D l = horiz ? new Line2D.Double(p, plot.getY(), p, plot.getMaxY()) : new Line2D.Double(plot.getX(), p, plot.getMaxX(), p);
                dp.strokeShape(cv, l, grid);
            }
            if (!valAx.deleted && !"none".equals(valAx.tickLblPos)) {
                String lab = axisLabel(valAx, v, srcFmt);
                if (horiz) text(cv, lab, p, plot.getMaxY() + 3, vt, 0, -1);
                else text(cv, lab, plot.getX() - 4, p, vt, 1, 0);
            }
        }
        if (scatter) {
            DrawingPainter.LineStyle xg = xValAx.gridMajor ? defaultLine(new Color(0xD9D9D9), 0.75) : null;
            for (double v = xValAx.lo; v <= xValAx.hi + xValAx.unit * 1e-6; v += xValAx.unit) {
                double p = plot.getX() + (v - xValAx.lo) / (xValAx.hi - xValAx.lo) * plot.getWidth();
                if (xg != null) dp.strokeShape(cv, new Line2D.Double(p, plot.getY(), p, plot.getMaxY()), xg);
                if (!xValAx.deleted) text(cv, axisLabel(xValAx, v, "General"), p, plot.getMaxY() + 3, ct, 0, -1);
            }
        }

        // 系列
        for (Group gg : groups) {
            switch (gg.type) {
                case "bar" -> paintBars(cv, gg, valAx, plot, nCats);
                case "line" -> paintLines(cv, gg, valAx, plot, nCats, false);
                case "area" -> paintLines(cv, gg, valAx, plot, nCats, true);
                case "scatter" -> paintScatter(cv, gg, valAx, xValAx, plot);
                default -> { }
            }
        }

        // 項目軸の線とラベル
        if (!scatter && catAx != null) {
            DrawingPainter.LineStyle axLine = catAx.el != null && Xml.child(catAx.el, "spPr") != null
                    ? dp.line(Xml.child(Xml.child(catAx.el, "spPr"), "ln"), null, null) : defaultLine(new Color(0xD9D9D9), 0.75);
            if (catAx.el != null && Xml.child(catAx.el, "spPr") != null && Xml.child(Xml.child(catAx.el, "spPr"), "ln") == null)
                axLine = defaultLine(new Color(0xD9D9D9), 0.75);
            double zero = valPos(valAx, Math.max(valAx.lo, Math.min(valAx.hi, 0)), plot, horiz);
            if (!catAx.deleted && axLine != null) {
                Line2D l = horiz ? new Line2D.Double(zero, plot.getY(), zero, plot.getMaxY()) : new Line2D.Double(plot.getX(), zero, plot.getMaxX(), zero);
                dp.strokeShape(cv, l, axLine);
            }
            if (!catAx.deleted && !"none".equals(catAx.tickLblPos)) {
                Series s0 = g.series.isEmpty() ? null : g.series.get(0);
                List<String> cats = s0 == null ? List.of() : s0.cats;
                for (int i = 0; i < nCats; i++) {
                    String lab = i < cats.size() ? cats.get(i) : String.valueOf(i + 1);
                    if (horiz) {
                        double cy = catCenter(i, nCats, plot, true, catAx.reverse);
                        text(cv, lab, plot.getX() - 4, cy, ct, 1, 0);
                    } else {
                        double cx = catCenter(i, nCats, plot, false, catAx.reverse);
                        double slot = plot.getWidth() / Math.max(1, nCats);
                        paintCatLabel(cv, lab, cx, plot.getMaxY() + 3, slot, ct);
                    }
                }
            }
        }
        if (valAx.titleEl != null) {
            TStyle tt = style(Xml.child(valAx.titleEl, "txPr"));
            tt.size = Math.max(tt.size, 10);
            String s = richText(valAx.titleEl);
            if (!s.isEmpty()) {
                cv.save();
                java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(box.getX() + lineH(tt) / 2, plot.getCenterY());
                at.rotate(-Math.PI / 2);
                cv.transform(at);
                text(cv, s, 0, 0, tt, 0, 0);
                cv.restore();
            }
        }
    }

    /** 書式付き文字列の文字だけ。 */
    private String richText(Element title) {
        Element rich = Xml.path(title, "tx", "rich");
        if (rich == null) return "";
        StringBuilder sb = new StringBuilder();
        for (Element p : Xml.children(rich, "p")) {
            for (Element r : Xml.children(p, "r")) sb.append(Xml.text(Xml.child(r, "t")));
        }
        return sb.toString();
    }

    /** 項目ラベルの高さ (折り返し込み)。 */
    private double catLabelHeight(Group g, Axis catAx, TStyle t, double slot) {
        Series s0 = g.series.isEmpty() ? null : g.series.get(0);
        int lines = 1;
        if (s0 != null) for (String c : s0.cats) {
            double w = textWidth(c, t);
            if (w > slot) lines = Math.max(lines, (int) Math.ceil(w / Math.max(slot, 1)));
        }
        return Math.min(lines, 3) * lineH(t);
    }

    /** 項目ラベルを描く (項目の幅で折り返す)。 */
    private void paintCatLabel(PdfCanvas cv, String lab, double cx, double y, double slot, TStyle t) throws IOException {
        double w = textWidth(lab, t);
        if (w <= slot + 1) {
            text(cv, lab, cx, y, t, 0, -1);
            return;
        }
        // 項目の幅に収まるよう折り返す
        TextLayout.Run run = new TextLayout.Run(lab, t.spec(), t.size, t.color);
        List<TextLayout.Line> lines = TextLayout.layout(List.of(run), (float) Math.max(slot, t.size), true, false, ctx.fonts);
        double yy = y;
        for (TextLayout.Line l : lines) {
            TextLayout.drawLine(cv, l, cx - l.width / 2, yy + l.ascent, 0, 0);
            yy += l.ascent + l.descent;
        }
    }

    /** 値の位置 (pt)。 */
    private static double valPos(Axis a, double v, Rectangle2D plot, boolean horiz) {
        double f = (v - a.lo) / (a.hi - a.lo);
        if (a.reverse) f = 1 - f;
        return horiz ? plot.getX() + f * plot.getWidth() : plot.getMaxY() - f * plot.getHeight();
    }

    /** 項目の中心の位置 (pt)。 */
    private static double catCenter(int i, int n, Rectangle2D plot, boolean horiz, boolean reverse) {
        double slot = (horiz ? plot.getHeight() : plot.getWidth()) / Math.max(1, n);
        int k = reverse ? n - 1 - i : i;
        if (horiz) return plot.getMaxY() - (k + 0.5) * slot; // 最初の項目は下
        return plot.getX() + (k + 0.5) * slot;
    }

    /** 棒グラフを描く。 */
    private void paintBars(PdfCanvas cv, Group g, Axis valAx, Rectangle2D plot, int n) throws IOException {
        boolean horiz = g.horizontal;
        boolean stacked = !g.grouping.equals("clustered") && !g.grouping.equals("standard");
        boolean percent = g.grouping.equals("percentStacked");
        int ns = Math.max(1, g.series.size());
        double slot = (horiz ? plot.getHeight() : plot.getWidth()) / Math.max(1, n);
        double gap = g.gapWidth / 100.0;
        double ov = g.overlap / 100.0;
        double barW = stacked ? slot / (1 + gap) : slot / (ns - (ns - 1) * ov + gap);
        double groupW = stacked ? barW : barW * (ns - (ns - 1) * ov);
        double zero = valPos(valAx, Math.max(valAx.lo, Math.min(valAx.hi, 0)), plot, horiz);
        double[] posAcc = new double[n], negAcc = new double[n], totals = new double[n];
        if (percent) for (Series s : g.series) for (int i = 0; i < n && i < s.vals.size(); i++) if (s.vals.get(i) != null) totals[i] += Math.abs(s.vals.get(i));
        for (int si = 0; si < g.series.size(); si++) {
            Series s = g.series.get(si);
            DrawingPainter.LineStyle ls = s.spPr == null ? null : dp.line(Xml.child(s.spPr, "ln"), null, null);
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                if (percent) v = totals[i] == 0 ? 0 : v / totals[i];
                double base;
                double from, to;
                if (stacked) {
                    base = v >= 0 ? posAcc[i] : negAcc[i];
                    if (v >= 0) posAcc[i] += v; else negAcc[i] += v;
                    from = valPos(valAx, clamp(base, valAx), plot, horiz);
                    to = valPos(valAx, clamp(base + v, valAx), plot, horiz);
                } else {
                    from = zero;
                    to = valPos(valAx, clamp(v, valAx), plot, horiz);
                }
                double c = catCenter(i, n, plot, horiz, false);
                double offset = stacked ? -barW / 2 : -groupW / 2 + si * barW * (1 - ov);
                Rectangle2D bar;
                if (horiz) {
                    double y0 = c + groupW / 2 - (stacked ? barW : (si + 1) * barW * (1 - ov) + barW * ov);
                    if (stacked) y0 = c - barW / 2;
                    bar = new Rectangle2D.Double(Math.min(from, to), y0, Math.abs(to - from), barW);
                } else {
                    bar = new Rectangle2D.Double(c + offset, Math.min(from, to), barW, Math.abs(to - from));
                }
                Color col = pointColor(g, s, i);
                DrawingPainter.Paint p = s.spPr != null && DrawingPainter.fillElement(s.spPr) != null
                        && !(g.varyColors && g.series.size() == 1) ? dp.paint(DrawingPainter.fillElement(s.spPr), null, null) : solid(col);
                if (p == null) p = solid(col);
                if (p.image == null && p.stopPos == null) p.color = col;
                dp.fillShape(cv, bar, p, null);
                if (ls != null) dp.strokeShape(cv, bar, ls);
                paintDataLabel(cv, g, s, i, v, horiz ? new double[] {Math.max(from, to), bar.getCenterY()} : new double[] {bar.getCenterX(), Math.min(from, to)}, bar);
            }
        }
    }

    /** 値を軸の範囲に収める。 */
    private static double clamp(double v, Axis a) { return Math.max(a.lo, Math.min(a.hi, v)); }

    /** 折れ線グラフ (area なら面グラフ) を描く。 */
    private void paintLines(PdfCanvas cv, Group g, Axis valAx, Rectangle2D plot, int n, boolean area) throws IOException {
        boolean stacked = g.grouping.equals("stacked") || g.grouping.equals("percentStacked");
        boolean percent = g.grouping.equals("percentStacked");
        double[] acc = new double[n], totals = new double[n];
        if (percent) for (Series s : g.series) for (int i = 0; i < n && i < s.vals.size(); i++) if (s.vals.get(i) != null) totals[i] += s.vals.get(i);
        // 積み上げたとき最初の系列が見えるよう、面は後ろの系列から描く
        List<double[][]> polys = new ArrayList<>();
        for (int si = 0; si < g.series.size(); si++) {
            Series s = g.series.get(si);
            double[][] pts = new double[n][];
            double[][] base = new double[n][];
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                if (percent) v = totals[i] == 0 ? 0 : v / totals[i];
                double b = stacked ? acc[i] : 0;
                double y = stacked ? b + v : v;
                if (stacked) acc[i] = y;
                double x = area && n > 1 ? plot.getX() + i * plot.getWidth() / (n - 1) : catCenter(i, n, plot, false, false);
                pts[i] = new double[] {x, valPos(valAx, clamp(y, valAx), plot, false)};
                base[i] = new double[] {x, valPos(valAx, clamp(b, valAx), plot, false)};
            }
            if (area) {
                Path2D p = new Path2D.Double();
                boolean started = false;
                for (double[] pt : pts) {
                    if (pt == null) continue;
                    if (!started) { p.moveTo(pt[0], pt[1]); started = true; } else p.lineTo(pt[0], pt[1]);
                }
                for (int i = n - 1; i >= 0; i--) if (base[i] != null) p.lineTo(base[i][0], base[i][1]);
                p.closePath();
                DrawingPainter.Paint fill = s.spPr != null && DrawingPainter.fillElement(s.spPr) != null
                        ? dp.paint(DrawingPainter.fillElement(s.spPr), null, null) : solid(s.color);
                dp.fillShape(cv, p, fill, null);
                DrawingPainter.LineStyle ls = s.spPr == null ? null : dp.line(Xml.child(s.spPr, "ln"), null, null);
                if (ls != null) dp.strokeShape(cv, p, ls);
            } else {
                polys.add(pts);
                DrawingPainter.LineStyle ls = seriesLine(s);
                Path2D p = new Path2D.Double();
                boolean started = false;
                for (double[] pt : pts) {
                    if (pt == null) { started = false; continue; }
                    if (!started) { p.moveTo(pt[0], pt[1]); started = true; } else p.lineTo(pt[0], pt[1]);
                }
                if (ls != null) dp.strokeShape(cv, p, ls);
                if (g.markers) for (int i = 0; i < n; i++) if (pts[i] != null) marker(cv, s, si, pts[i][0], pts[i][1]);
            }
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                if (pts[i] != null && s.vals.get(i) != null)
                    paintDataLabel(cv, g, s, i, s.vals.get(i), pts[i], null);
            }
        }
    }

    /** 系列の線のスタイル。 */
    private DrawingPainter.LineStyle seriesLine(Series s) {
        DrawingPainter.LineStyle ls = null;
        if (s.spPr != null && Xml.child(s.spPr, "ln") != null) ls = dp.line(Xml.child(s.spPr, "ln"), null, null);
        else {
            ls = defaultLine(s.color, 2.25);
            ls.cap = 1;
            ls.join = 1;
        }
        if (ls != null && ls.color == null && s.spPr != null && Xml.child(Xml.child(s.spPr, "ln"), "noFill") == null) ls.color = s.color;
        return ls;
    }

    /** マーカーを描く。 */
    private void marker(PdfCanvas cv, Series s, int si, double x, double y) throws IOException {
        String sym = "auto";
        double size = 5;
        Color fill = s.color, line = s.color;
        if (s.marker != null) {
            sym = Xml.attr(Xml.child(s.marker, "symbol"), "val", "auto");
            size = Xml.doubleAttr(Xml.child(s.marker, "size"), "val", 5);
            Element spPr = Xml.child(s.marker, "spPr");
            if (spPr != null) {
                Element f = DrawingPainter.fillElement(spPr);
                if (f != null) {
                    DrawingPainter.Paint p = dp.paint(f, null, null);
                    fill = p == null ? null : p.color;
                }
                DrawingPainter.LineStyle l = dp.line(Xml.child(spPr, "ln"), null, null);
                if (Xml.child(spPr, "ln") != null) line = l == null ? null : l.color;
            }
        }
        if (sym.equals("none")) return;
        if (sym.equals("auto")) {
            String[] auto = {"diamond", "square", "triangle", "x", "star", "circle", "plus", "dash"};
            sym = auto[si % auto.length];
        }
        double h = size * 0.75 / 2 + 1;
        Shape sh = switch (sym) {
            case "circle" -> new Ellipse2D.Double(x - h, y - h, 2 * h, 2 * h);
            case "diamond" -> poly(x, y - h, x + h, y, x, y + h, x - h, y);
            case "triangle" -> poly(x, y - h, x + h, y + h, x - h, y + h);
            default -> new Rectangle2D.Double(x - h, y - h, 2 * h, 2 * h);
        };
        if (sym.equals("x") || sym.equals("plus") || sym.equals("star") || sym.equals("dash")) {
            PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.75f);
            Color c = line != null ? line : s.color;
            if (!sym.equals("plus") && !sym.equals("dash")) {
                cv.line(x - h, y - h, x + h, y + h, st, c);
                cv.line(x - h, y + h, x + h, y - h, st, c);
            }
            if (!sym.equals("x")) cv.line(x - h, y, x + h, y, st, c);
            if (sym.equals("plus") || sym.equals("star")) cv.line(x, y - h, x, y + h, st, c);
            return;
        }
        if (fill != null) cv.fill(sh, fill);
        if (line != null) cv.stroke(sh, new PdfCanvas.Stroke(0.75f), line);
    }

    /** 座標の並びから閉じた多角形を作る。 */
    private static Path2D poly(double... c) {
        Path2D p = new Path2D.Double();
        p.moveTo(c[0], c[1]);
        for (int i = 2; i < c.length; i += 2) p.lineTo(c[i], c[i + 1]);
        p.closePath();
        return p;
    }

    /** 散布図を描く。 */
    private void paintScatter(PdfCanvas cv, Group g, Axis yAx, Axis xAx, Rectangle2D plot) throws IOException {
        boolean lines = !g.scatterStyle.equals("marker");
        for (int si = 0; si < g.series.size(); si++) {
            Series s = g.series.get(si);
            List<double[]> pts = new ArrayList<>();
            for (int i = 0; i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                double xv = i < s.xs.size() && s.xs.get(i) != null ? s.xs.get(i) : i + 1;
                double x = plot.getX() + (xv - xAx.lo) / (xAx.hi - xAx.lo) * plot.getWidth();
                pts.add(new double[] {x, valPos(yAx, v, plot, false)});
            }
            boolean hasLine = s.spPr != null && Xml.child(s.spPr, "ln") != null && Xml.child(Xml.child(s.spPr, "ln"), "noFill") == null;
            if ((lines && (s.spPr == null || Xml.child(s.spPr, "ln") == null)) || hasLine) {
                Path2D p = new Path2D.Double();
                for (int i = 0; i < pts.size(); i++) {
                    if (i == 0) p.moveTo(pts.get(i)[0], pts.get(i)[1]); else p.lineTo(pts.get(i)[0], pts.get(i)[1]);
                }
                dp.strokeShape(cv, p, seriesLine(s));
            }
            for (double[] pt : pts) marker(cv, s, si, pt[0], pt[1]);
        }
    }

    /** 円グラフ・ドーナツグラフを描く。 */
    private void paintPie(PdfCanvas cv, Group g, Rectangle2D box) throws IOException {
        if (g.series.isEmpty()) return;
        Series s = g.series.get(0);
        double total = 0;
        for (Double v : s.vals) if (v != null && v > 0) total += v;
        if (total <= 0) return;
        double d = Math.min(box.getWidth(), box.getHeight()) * 0.9;
        double cx = box.getCenterX(), cy = box.getCenterY();
        Rectangle2D circle = new Rectangle2D.Double(cx - d / 2, cy - d / 2, d, d);
        double start = 90 - g.firstSliceAng; // Java の角度は 3 時の方向から反時計回り
        double hole = g.type.equals("doughnut") ? g.holeSize / 100.0 : 0;
        DrawingPainter.LineStyle edge = s.spPr != null && Xml.child(s.spPr, "ln") != null ? dp.line(Xml.child(s.spPr, "ln"), null, null)
                : defaultLine(Color.WHITE, 0.75);
        for (int i = 0; i < s.vals.size(); i++) {
            Double v = s.vals.get(i);
            if (v == null || v <= 0) continue;
            double ext = -v / total * 360;
            Shape slice;
            if (hole > 0) {
                Arc2D outer = new Arc2D.Double(circle, start, ext, Arc2D.OPEN);
                double hd = d * hole;
                Arc2D innerArc = new Arc2D.Double(cx - hd / 2, cy - hd / 2, hd, hd, start + ext, -ext, Arc2D.OPEN);
                Path2D p = new Path2D.Double(outer);
                p.append(innerArc, true);
                p.closePath();
                slice = p;
            } else {
                slice = new Arc2D.Double(circle, start, ext, Arc2D.PIE);
            }
            Color c = pointColor(g, s, i);
            cv.fill(slice, c);
            if (edge != null) dp.strokeShape(cv, slice, edge);
            double mid = Math.toRadians(start + ext / 2);
            double lr = d / 2 * (hole > 0 ? (1 + hole) / 2 : 0.6);
            paintDataLabel(cv, g, s, i, v, new double[] {cx + Math.cos(mid) * lr, cy - Math.sin(mid) * lr}, null, v / total);
            start += ext;
        }
    }

    /** レーダーチャートを描く。 */
    private void paintRadar(PdfCanvas cv, Group g, List<Axis> axes, Rectangle2D box) throws IOException {
        int n = 0;
        double max = 0;
        for (Series s : g.series) {
            n = Math.max(n, s.vals.size());
            for (Double v : s.vals) if (v != null) max = Math.max(max, v);
        }
        if (n < 3) return;
        Axis va = new Axis();
        for (Axis a : axes) if (a.kind.equals("valAx")) va = a;
        autoScale(va, 0, max, Math.min(box.getWidth(), box.getHeight()) / 2);
        double rad = Math.min(box.getWidth(), box.getHeight()) / 2 * 0.8;
        double cx = box.getCenterX(), cy = box.getCenterY();
        DrawingPainter.LineStyle grid = defaultLine(new Color(0xD9D9D9), 0.75);
        for (double v = va.lo; v <= va.hi + 1e-9; v += va.unit) {
            double rr = (v - va.lo) / (va.hi - va.lo) * rad;
            Path2D p = new Path2D.Double();
            for (int i = 0; i <= n; i++) {
                double a = Math.PI / 2 - 2 * Math.PI * i / n;
                double x = cx + Math.cos(a) * rr, y = cy - Math.sin(a) * rr;
                if (i == 0) p.moveTo(x, y); else p.lineTo(x, y);
            }
            dp.strokeShape(cv, p, grid);
        }
        for (int si = 0; si < g.series.size(); si++) {
            Series s = g.series.get(si);
            Path2D p = new Path2D.Double();
            for (int i = 0; i <= n; i++) {
                Double v = s.vals.get(i % n);
                double rr = ((v == null ? 0 : v) - va.lo) / (va.hi - va.lo) * rad;
                double a = Math.PI / 2 - 2 * Math.PI * i / n;
                double x = cx + Math.cos(a) * rr, y = cy - Math.sin(a) * rr;
                if (i == 0) p.moveTo(x, y); else p.lineTo(x, y);
            }
            if (g.radarStyle.equals("filled")) cv.fill(p, s.color);
            else dp.strokeShape(cv, p, seriesLine(s));
        }
    }

    // ------------------------------------------------------------------ データラベル

    /** データラベルを描く。 */
    private void paintDataLabel(PdfCanvas cv, Group g, Series s, int i, double v, double[] at, Rectangle2D bar) throws IOException {
        paintDataLabel(cv, g, s, i, v, at, bar, Double.NaN);
    }

    /**
     * データラベルを描く。
     *
     * @param at  ラベルの基準位置 {x, y}
     * @param bar 棒の矩形 (棒グラフのとき。なければ null)
     * @param pct 円グラフの割合 (なければ NaN)
     */
    private void paintDataLabel(PdfCanvas cv, Group g, Series s, int i, double v, double[] at, Rectangle2D bar, double pct)
            throws IOException {
        Element dl = s.dLbls != null ? s.dLbls : g.dLbls;
        if (dl == null) return;
        if (Xml.boolAttr(Xml.child(dl, "delete"), "val", false)) return;
        boolean showVal = Xml.boolAttr(Xml.child(dl, "showVal"), "val", false);
        boolean showCat = Xml.boolAttr(Xml.child(dl, "showCatName"), "val", false);
        boolean showSer = Xml.boolAttr(Xml.child(dl, "showSerName"), "val", false);
        boolean showPct = Xml.boolAttr(Xml.child(dl, "showPercent"), "val", false);
        // データ要素ごとの指定
        for (Element d : Xml.children(dl, "dLbl")) {
            if (Xml.intAttr(Xml.child(d, "idx"), "val", -1) == i) {
                if (Xml.boolAttr(Xml.child(d, "delete"), "val", false)) return;
                if (Xml.child(d, "showVal") != null) showVal = Xml.boolAttr(Xml.child(d, "showVal"), "val", false);
            }
        }
        if (!showVal && !showCat && !showSer && !showPct) return;
        List<String> parts = new ArrayList<>();
        if (showSer) parts.add(s.name);
        if (showCat && i < s.cats.size()) parts.add(s.cats.get(i));
        Element nf = Xml.child(dl, "numFmt");
        String fmt = nf != null && !Xml.boolAttr(nf, "sourceLinked", false) ? nf.getAttribute("formatCode") : s.fmt;
        if (showVal) parts.add(ctx.formatter.formatNumber(v, fmt));
        if (showPct && !Double.isNaN(pct)) parts.add(ctx.formatter.formatNumber(pct, "0%"));
        String lab = String.join(Xml.text(Xml.child(dl, "separator")).isEmpty() ? (showPct && showCat ? "\n" : ", ") : Xml.text(Xml.child(dl, "separator")), parts);
        TStyle t = style(Xml.child(dl, "txPr"));
        String pos = Xml.attr(Xml.child(dl, "dLblPos"), "val", bar != null ? "outEnd" : g.type.equals("pie") ? "bestFit" : "r");
        double x = at[0], y = at[1];
        for (String line : lab.split("\n")) {
            if (bar != null && !g.horizontal) {
                switch (pos) {
                    case "ctr" -> text(cv, line, bar.getCenterX(), bar.getCenterY(), t, 0, 0);
                    case "inEnd" -> text(cv, line, x, y + 2, t, 0, -1);
                    case "inBase" -> text(cv, line, x, bar.getMaxY() - 2, t, 0, 1);
                    default -> text(cv, line, x, y - 2, t, 0, 1);
                }
            } else if (bar != null) {
                switch (pos) {
                    case "ctr" -> text(cv, line, bar.getCenterX(), bar.getCenterY(), t, 0, 0);
                    default -> text(cv, line, x + 3, y, t, -1, 0);
                }
            } else if (g.type.equals("pie") || g.type.equals("doughnut")) {
                text(cv, line, x, y, t, 0, 0);
                y += lineH(t);
            } else {
                switch (pos) {
                    case "t" -> text(cv, line, x, y - 4, t, 0, 1);
                    case "b" -> text(cv, line, x, y + 4, t, 0, -1);
                    case "l" -> text(cv, line, x - 5, y, t, 1, 0);
                    case "ctr" -> text(cv, line, x, y, t, 0, 0);
                    default -> text(cv, line, x + 5, y, t, -1, 0);
                }
            }
        }
    }

    /** 数値を小数 2 桁の文字列にする (デバッグ用)。 */
    static String fmt(double d) {
        return String.format(Locale.ROOT, "%.2f", d);
    }
}
