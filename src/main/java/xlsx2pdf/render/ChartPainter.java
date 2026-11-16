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
    /** グラフ全体に書式がなく、旧来のスタイルの既定で描く。 */
    private boolean legacy;
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
        /** 系列名がファイルにある (なければ「系列 N」)。 */
        boolean named;
        /** バブルの大きさ。 */
        List<Double> sizes = new ArrayList<>();
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
        /** 等高線グラフをワイヤーフレームで描く。 */
        boolean wireframe;
        /** 3D のグラフと、系列の奥行きの間隔 (%)。 */
        boolean is3D;
        double gapDepth = 150;
        /** 補助円グラフの種類 (pie / bar。補助円でなければ null)、補助に回す最後の要素の数、補助の大きさ (%)、区分線。 */
        String ofPie;
        int splitPos = 2;
        double secondPieSize = 75;
        Element serLines;
        /** 3D の棒の形 (box, cylinder, cone, coneToMax, pyramid, pyramidToMax)。 */
        String shape = "box";
        /** バブルチャートと、バブルの大きさの倍率 (%)。 */
        boolean bubble;
        double bubbleScale = 100;
        /** バブルの大きさが幅 (直径) を表す (sizeRepresents="w")。既定は面積。 */
        boolean bubbleWidth;
        /** 高低線、ローソク (陽線・陰線)、降下線 (なければ null)。 */
        Element hiLowLines, upDownBars, dropLines;
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
        /** 上端・下端の余裕を足さない (3D のグラフ。Excel の印刷: 積み上げの最大 20 で 0〜20)。 */
        boolean noHeadroom;
        Element gridMinorSpPr;
        /** 補助目盛の間隔 (null は主目盛の 1/5)。 */
        Double minorUnit;
        Element gridSpPr;
        /** 目盛ラベルの位置。 */
        String tickLblPos = "nextTo";
        /** 主目盛の種類 (none, out, in, cross)。 */
        String majorTick = "none";
        /** 交差する位置と、項目の境界での交差。 */
        String crosses = "autoZero";
        String crossBetween = "between";
        /** 交差する位置の値 (crossesAt。なければ null)。 */
        Double crossesAt;
        /** 交差する軸の ID。 */
        String crossAx;
        /** 計算した目盛りの範囲と間隔。 */
        double lo, hi, unit;
        /** 対数目盛。 */
        boolean logScale;
        /** 表示単位 (目盛ラベルを割る値。1 は単位なし) とラベルを表示するか。 */
        double dispUnit = 1;
        String dispUnitName;
        boolean dispUnitLbl;
        /** 上端の余裕の目盛りを、最大値が目盛りにちょうど重なるときだけ足す (Excel 2016 以降のグラフ。exp90 の箱ひげ図)。 */
        boolean exactHeadroom;
        /** 目盛りの区間の数の上限。 */
        int maxIntervals = MAX_INTERVALS;
        /** 対数目盛の底。 */
        double logBase = 10;
        /** 軸ラベル。 */
        String title;
        Element titleEl;
    }

    // ------------------------------------------------------------------ 描画

    /** グラフエリアの外枠のずれ (pt)。 */
    static final double FRAME_DX = -0.48, FRAME_DY = -0.6;

    /** グラフを矩形 r に描く: グラフエリア、タイトル、凡例、プロットエリア。 */
    public void paint(PdfCanvas cv, Rectangle2D r) throws IOException {
        Element root;
        try (InputStream in = part.getInputStream()) {
            root = Xml.parse(in);
        } catch (Exception e) {
            Warnings.add("グラフの定義を読めませんでした", e);
            return;
        }
        defaultTxPr = Xml.child(root, "txPr");
        dTable = Xml.path(root, "chart", "plotArea", "dTable");
        // 空のセルの表示 (gap = 線を切る、zero = 0、span = 線でつなぐ)。要素がなければ gap
        Element dba = Xml.child(Xml.child(root, "chart"), "dispBlanksAs");
        blanks = dba == null ? "gap" : Xml.attr(dba, "val", "zero");
        // グラフ全体に書式 (spPr) がない: プログラムで作ったブックなど。Excel は旧来のスタイル (c:style) の既定で描く
        legacy = Xml.child(root, "spPr") == null;
        Element chart = Xml.child(root, "chart");
        if (chart == null) return;
        // Excel のグラフエリアはアンカーより左上に寄る (外枠の実測 0.48pt・0.6pt)。
        // 新しいスタイルの中身はアンカー基準で合わせ込んであるので外枠と背景だけ動かし、旧来のスタイルとレーダーは中身も動かす
        Rectangle2D frame = new Rectangle2D.Double(r.getX() + FRAME_DX, r.getY() + FRAME_DY, r.getWidth(), r.getHeight());
        if (legacy || Xml.path(root, "chart", "plotArea", "radarChart") != null) r = frame;
        chartEl = chart;
        boolean rounded = Xml.boolAttr(Xml.child(root, "roundedCorners"), "val", false);
        if (Xml.child(root, "roundedCorners") == null) rounded = false;

        // グラフエリアの背景
        chartRect = r;
        dumpPlot = null; dumpLegend = null; dumpVal = null;
        Element v3 = Xml.path(root, "chart", "view3D");
        rotX = Xml.doubleAttr(Xml.child(v3, "rotX"), "val", 15);
        rotY = Xml.doubleAttr(Xml.child(v3, "rotY"), "val", 20);
        depthPct = Xml.doubleAttr(Xml.child(v3, "depthPercent"), "val", 100);
        rAngAx = Xml.boolAttr(Xml.child(v3, "rAngAx"), "val", true);
        perspective = Xml.doubleAttr(Xml.child(v3, "perspective"), "val", 30);
        hPercent = Xml.child(v3, "hPercent") == null ? null : Xml.doubleAttr(Xml.child(v3, "hPercent"), "val", 100);
        Element csSpPr = Xml.child(root, "spPr");
        Shape area = rounded ? new java.awt.geom.RoundRectangle2D.Double(r.getX(), r.getY(), r.getWidth(), r.getHeight(), 10, 10) : r;
        Shape frameArea = rounded ? new java.awt.geom.RoundRectangle2D.Double(frame.getX(), frame.getY(), frame.getWidth(), frame.getHeight(), 10, 10) : frame;
        DrawingPainter.Paint bg = csSpPr != null ? dp.paint(DrawingPainter.fillElement(csSpPr), null, null) : solid(Color.WHITE);
        if (csSpPr != null && DrawingPainter.fillElement(csSpPr) == null) bg = solid(Color.WHITE);
        dp.fillShape(cv, frameArea, bg, null);
        DrawingPainter.LineStyle border = csSpPr != null ? dp.line(Xml.child(csSpPr, "ln"), null, null) : null;
        if (csSpPr == null || Xml.child(csSpPr, "ln") == null) border = gridDefault(); // 旧来のスタイルも目盛線と同じ灰色の線 (exp75)

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
        titleExtra = 0;
        Element title = Xml.child(chart, "title");
        boolean autoDeleted = Xml.boolAttr(Xml.child(chart, "autoTitleDeleted"), "val", false);
        String autoTitle = null;
        if (title == null && !autoDeleted && groups.size() == 1 && groups.get(0).series.size() == 1
                && (groups.get(0).type.equals("pie") || groups.get(0).type.equals("doughnut"))) {
            autoTitle = groups.get(0).series.get(0).name;
        }
        if (title != null || autoTitle != null) {
            String fallback = title != null ? null : autoTitle;
            // 系列名を題名にするのは、グラフ全体で系列が 1 つのとき (組み合わせグラフで 2 つなら「グラフ タイトル」。exp72)
            // 円グラフは先頭の系列だけを描くので、系列が複数でも先頭の系列名 (rand30 rd08)
            if (title != null && Xml.child(title, "tx") == null && groups.size() > 0
                    && (allSeries(groups).size() == 1 || groups.size() == 1 && groups.get(0).type.equals("pie")))
                fallback = groups.get(0).series.get(0).named ? groups.get(0).series.get(0).name : null;
            if (title != null && Xml.child(title, "tx") == null && fallback == null) fallback = "グラフ タイトル";
            double h = paintTitle(cv, title, fallback, inner, legacy ? 18 : 14, true);
            boolean overlay = title != null && Xml.boolAttr(Xml.child(title, "overlay"), "val", false);
            if (!overlay) top += h + 4 - (legacy ? LEGACY_TITLE_CUT : 0);
        }

        // 凡例
        Element legend = Xml.child(chart, "legend");
        Rectangle2D plotBox = new Rectangle2D.Double(inner.getX(), top, inner.getWidth(), inner.getMaxY() - top);
        chartAxes = axes;
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
            if (border != null) dp.strokeShape(cv, frameArea, border);
            dumpChart(r);
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
        if (border != null) dp.strokeShape(cv, frameArea, border);
        dumpChart(r);
    }

    /** 値の書き出し (開発用): グラフの位置、プロットの内側・凡例はグラフの左上からの pt、値の軸の範囲 (scripts/dump-excel.ps1 と同じ形)。 */
    private Rectangle2D dumpPlot, dumpLegend;
    private Axis dumpVal;

    private void dumpChart(Rectangle2D r) {
        if (!ValueDump.ON) return;
        ValueDump.add("chart", "-", r.getX(), r.getY(), r.getWidth(), r.getHeight(), "-");
        if (dumpPlot != null)
            ValueDump.add("plot", "-", dumpPlot.getX() - r.getX(), dumpPlot.getY() - r.getY(), dumpPlot.getWidth(), dumpPlot.getHeight());
        if (dumpLegend != null)
            ValueDump.add("legend", "-", dumpLegend.getX() - r.getX(), dumpLegend.getY() - r.getY(), dumpLegend.getWidth(), dumpLegend.getHeight());
        if (dumpVal != null) ValueDump.add("vaxis", "-", dumpVal.lo, dumpVal.hi, dumpVal.unit);
    }

    /** 単色の塗りつぶし。 */
    private static DrawingPainter.Paint solid(Color c) {
        DrawingPainter.Paint p = new DrawingPainter.Paint();
        p.color = c;
        return p;
    }

    /** 目盛線・軸の線の既定 (旧来のスタイルは灰色 1pt、2013 以降は薄い灰色 0.75pt)。 */
    private DrawingPainter.LineStyle gridDefault() {
        return legacy ? defaultLine(new Color(0x898989), 1) : defaultLine(new Color(0xD9D9D9), 0.75);
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
                g.serLines = Xml.child(e, "serLines");
                g.shape = Xml.attr(Xml.child(e, "shape"), "val", "box");
            }
            case "lineChart", "line3DChart", "stockChart" -> g.type = "line";
            case "areaChart", "area3DChart" -> g.type = "area";
            case "pieChart", "pie3DChart" -> g.type = "pie";
            case "ofPieChart" -> {
                g.type = "pie";
                g.ofPie = Xml.attr(Xml.child(e, "ofPieType"), "val", "pie");
                g.splitPos = Xml.intAttr(Xml.child(e, "splitPos"), "val", 2);
                g.secondPieSize = Xml.doubleAttr(Xml.child(e, "secondPieSize"), "val", 75);
                g.gapWidth = Xml.doubleAttr(Xml.child(e, "gapWidth"), "val", 100);
                g.serLines = Xml.child(e, "serLines");
            }
            case "doughnutChart" -> {
                g.type = "doughnut";
                g.holeSize = Xml.doubleAttr(Xml.child(e, "holeSize"), "val", 50);
            }
            case "scatterChart", "bubbleChart" -> {
                g.type = "scatter";
                g.bubble = n.equals("bubbleChart");
                g.bubbleScale = Xml.doubleAttr(Xml.child(e, "bubbleScale"), "val", 100);
                g.bubbleWidth = "w".equals(Xml.attr(Xml.child(e, "sizeRepresents"), "val", "area"));
                g.scatterStyle = Xml.attr(Xml.child(e, "scatterStyle"), "val", "lineMarker");
            }
            case "surfaceChart", "surface3DChart" -> {
                g.type = "surface";
                g.wireframe = Xml.boolAttr(Xml.child(e, "wireframe"), "val", false);
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
        g.is3D = n.endsWith("3DChart");
        g.gapDepth = Xml.doubleAttr(Xml.child(e, "gapDepth"), "val", 150);
        g.hiLowLines = Xml.child(e, "hiLowLines");
        g.upDownBars = Xml.child(e, "upDownBars");
        g.dropLines = Xml.child(e, "dropLines");
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
        se.named = !se.name.isEmpty();
        if (se.name.isEmpty()) se.name = "系列" + (se.idx + 1);
        Element bs = Xml.child(s, "bubbleSize");
        if (bs != null) se.sizes = numbers(Xml.firstChild(bs));
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
        if (ctx.recalc && !cache.getLocalName().equals("multiLvlStrCache")) {
            List<org.apache.poi.ss.usermodel.Cell> cells = refCells(ref, pointCount(cache));
            if (cells != null) {
                for (org.apache.poi.ss.usermodel.Cell c : cells)
                    out.add(c == null ? "" : ctx.formatter.format(c, true, null).text);
                return out;
            }
        }
        Element lvlParent = cache;
        if (cache.getLocalName().equals("multiLvlStrCache")) {
            List<Element> lvls = Xml.children(cache, "lvl");
            if (lvls.isEmpty()) return out;
            lvlParent = lvls.get(0);
        }
        int count = pointCount(cache);
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

    /** グラフのデータの件数の上限 (Excel のシートの行数)。壊れた・不正なファイルで巨大な配列を確保しない。 */
    static final int MAX_POINTS = 1_048_576;

    /** キャッシュの件数 (ptCount)。負の値は 0、上限を超える値は上限にする。 */
    private static int pointCount(Element cache) {
        return Math.max(0, Math.min(Xml.intAttr(Xml.child(cache, "ptCount"), "val", 0), MAX_POINTS));
    }

    /** 数値の参照またはリテラルの値 (数値でなければ null)。 */
    private List<Double> numbers(Element ref) {
        List<Double> out = new ArrayList<>();
        if (ref == null) return out;
        Element cache = Xml.child(ref, "numCache");
        if (cache == null && ref.getLocalName().equals("numLit")) cache = ref;
        if (cache == null) return out;
        int count = pointCount(cache);
        if (ctx.recalc) {
            List<org.apache.poi.ss.usermodel.Cell> cells = refCells(ref, count);
            if (cells != null) {
                for (org.apache.poi.ss.usermodel.Cell c : cells) out.add(cellNumber(c));
                return out;
            }
        }
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

    /**
     * --recalc: 系列の参照 (numRef / strRef の f) の先のセルを、並びの順に返す。参照が 1 行か 1 列の連続した範囲で、
     * セルの数がキャッシュの件数 (ptCount) と同じときだけ使う (非表示の行・列を除いたキャッシュなどは、保存された値のまま)。
     *
     * @return セルの一覧 (セルのない位置は null)。使えなければ null
     */
    private List<org.apache.poi.ss.usermodel.Cell> refCells(Element ref, int count) {
        String f = Xml.text(Xml.child(ref, "f")).strip();
        if (f.isEmpty() || f.startsWith("(") || f.contains(",")) return null;
        try {
            var ar = new org.apache.poi.ss.util.AreaReference(f, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
            var first = ar.getFirstCell();
            var last = ar.getLastCell();
            String sheetName = first.getSheetName();
            if (sheetName == null) return null;
            var sh = ctx.wb.getSheet(sheetName);
            if (sh == null) return null;
            int r0 = first.getRow(), r1 = last.getRow(), c0 = first.getCol(), c1 = last.getCol();
            if (r0 != r1 && c0 != c1) return null;
            int n = (r1 - r0 + 1) * (c1 - c0 + 1);
            if (n != count || n > MAX_POINTS) return null;
            List<org.apache.poi.ss.usermodel.Cell> out = new ArrayList<>(n);
            for (int r = r0; r <= r1; r++) {
                var row = sh.getRow(r);
                for (int c = c0; c <= c1; c++) out.add(row == null ? null : row.getCell(c));
            }
            return out;
        } catch (RuntimeException e) {
            return null;
        }
    }

    /** セルの数値 (数値でなければ null。数式は計算結果)。 */
    private static Double cellNumber(org.apache.poi.ss.usermodel.Cell c) {
        if (c == null) return null;
        var t = c.getCellType();
        if (t == org.apache.poi.ss.usermodel.CellType.FORMULA) t = c.getCachedFormulaResultType();
        return t == org.apache.poi.ss.usermodel.CellType.NUMERIC ? c.getNumericCellValue() : null;
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
            if (a.logScale) a.logBase = Math.max(2, Xml.doubleAttr(Xml.child(sc, "logBase"), "val", 10));
        }
        if (Xml.child(e, "majorUnit") != null) a.majorUnit = Xml.doubleAttr(Xml.child(e, "majorUnit"), "val", 1);
        Element nf = Xml.child(e, "numFmt");
        if (nf != null && !Xml.boolAttr(nf, "sourceLinked", false)) a.fmt = nf.getAttribute("formatCode");
        Element g = Xml.child(e, "majorGridlines");
        a.gridMajor = g != null;
        a.gridSpPr = g == null ? null : Xml.child(g, "spPr");
        a.gridMinor = Xml.child(e, "minorGridlines") != null;
        a.gridMinorSpPr = a.gridMinor ? Xml.child(Xml.child(e, "minorGridlines"), "spPr") : null;
        if (Xml.child(e, "minorUnit") != null) a.minorUnit = Xml.doubleAttr(Xml.child(e, "minorUnit"), "val", 0);
        a.tickLblPos = Xml.attr(Xml.child(e, "tickLblPos"), "val", "nextTo");
        a.majorTick = Xml.attr(Xml.child(e, "majorTickMark"), "val", "cross");
        a.crosses = Xml.attr(Xml.child(e, "crosses"), "val", "autoZero");
        if (Xml.child(e, "crossesAt") != null) a.crossesAt = Xml.doubleAttr(Xml.child(e, "crossesAt"), "val", 0);
        a.crossBetween = Xml.attr(Xml.child(e, "crossBetween"), "val", "between");
        a.crossAx = Xml.attr(Xml.child(e, "crossAx"), "val", "");
        a.titleEl = Xml.child(e, "title");
        Element du = Xml.child(e, "dispUnits");
        if (du != null) {
            Element bu = Xml.child(du, "builtInUnit"), cu = Xml.child(du, "custUnit");
            if (bu != null) {
                String u = Xml.attr(bu, "val", "thousands");
                a.dispUnitName = u;
                a.dispUnit = switch (u) {
                    case "hundreds" -> 1e2; case "thousands" -> 1e3; case "tenThousands" -> 1e4; case "hundredThousands" -> 1e5;
                    case "millions" -> 1e6; case "tenMillions" -> 1e7; case "hundredMillions" -> 1e8; case "billions" -> 1e9;
                    case "trillions" -> 1e12; default -> 1;
                };
            } else if (cu != null) a.dispUnit = Xml.doubleAttr(cu, "val", 1);
            if (a.dispUnit <= 0) a.dispUnit = 1;
            a.dispUnitLbl = Xml.child(du, "dispUnitsLbl") != null;
        }
        return a;
    }

    // ------------------------------------------------------------------ 色と文字

    /** 系列の色 (明示の指定、なければ既定の色)。 */
    /** 白黒印刷の面の枠の太さ (pt。exp124 bw_area) */
    static final double BW_AREA_LINE = 1.5;
    /** 白黒印刷の系列・要素の色の順 (Excel の印刷: 黒、白、灰色の模様に近い濃さ)。 */
    private static final Color[] BW = {Color.BLACK, Color.WHITE, new Color(0x808080), new Color(0x404040), new Color(0xC0C0C0),
            new Color(0xE0E0E0), new Color(0x909090), new Color(0x707070)};
    /** 白黒印刷の 3 番目からの網かけ (Excel の印刷 exp126: 点の濃さ 4 つと横縞・縦縞。色の近似は BW の同じ位置) */
    private static final String[] BW_PATT = {null, null, "pct50", "pct75", "pct25", "pct10", "ltHorz", "ltVert"};

    /** 塗りの色。白黒印刷の網かけの色なら網かけにする。 */
    private DrawingPainter.Paint bwFill(Color c) {
        if (dp.bw && c != null) for (int k = 2; k < BW.length; k++) if (BW[k].equals(c)) return DrawingPainter.pattern(BW_PATT[k], Color.BLACK, Color.WHITE, c);
        return solid(c);
    }

    private Color seriesColor(Series s, int idx, int pos) {
        if (dp.bw) return BW[pos % BW.length];
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

    /**
     * 散布図の軸 a が、もう一方の軸 other の 0 の位置 (プロットの途中) で交わるか (a の crosses が既定の autoZero で、other の範囲が 0 をまたぐ)。
     */
    private static boolean scatterZeroInside(boolean scatter, Axis a, Axis other) {
        if (!scatter || a == null || other == null) return false;
        if ("max".equals(a.crosses) || "min".equals(a.crosses) || a.crossesAt != null) return false;
        return other.lo < -1e-9 && other.hi > 1e-9 && !other.logScale;
    }

    /** 系列の線が「なし」。 */
    private static boolean lineNone(Series s) {
        Element ln = s.spPr == null ? null : Xml.child(s.spPr, "ln");
        return ln != null && Xml.child(ln, "noFill") != null;
    }

    /** 系列の塗りが「塗りなし」。 */
    private static boolean noFill(Series s) {
        Element fe = s.spPr == null ? null : DrawingPainter.fillElement(s.spPr);
        return fe != null && "noFill".equals(fe.getLocalName());
    }

    /** データ要素 i の個別の書式 (dPt の spPr)。なければ null。 */
    private static Element pointSpPr(Series s, int i) {
        for (Element dpt : Xml.children(s.ser, "dPt"))
            if (Xml.intAttr(Xml.child(dpt, "idx"), "val", -1) == i) return Xml.child(dpt, "spPr");
        return null;
    }

    /** データ要素の色 (要素ごとの指定、要素ごとに色を変える設定、系列の色の順)。 */
    private Color pointColor(Group g, Series s, int i) {
        if (dp.bw) return g.varyColors && g.series.size() == 1 ? BW[i % BW.length] : s.color;
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
        if (dp.bw) return g.varyColors && g.series.size() == 1 ? BW[i % BW.length] : s.color;
        if (g.varyColors && g.series.size() == 1) {
            if (legacy && (g.type.equals("pie") || g.type.equals("doughnut"))) {
                // 旧来のスタイルの円 (Excel の印刷 exp75): 色 × 0.88。7〜12 番目は accent を白と半々に混ぜた色 × 0.88
                Color c = (i / 6) % 2 == 1 ? wireColor(autoColor(i % 6)) : autoColor(i);
                return shade(c, LEGACY_PIE_SHADE);
            }
            return autoColor(i);
        }
        return s.color;
    }

    /** txPr の連なりから決まる文字の書式。 */
    private final class TStyle {
        /** サイズ (pt)。 */
        float size = 10;
        /** 太字。 */
        boolean bold;
        /** 色 (既定は Excel のグラフの文字の灰色)。 */
        Color color = legacy ? Color.BLACK : new Color(0x595959);
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
        // グラフの文字の行の高さ: 行の高さの表の値と、東アジアのフォントの winAscent + winDescent の 1.2 倍の大きい方
        // (游ゴシックは表の 1.565em、ＭＳ Ｐゴシックは 1.2em、メイリオは 1.8em。Excel の印刷 exp92)
        FontManager.Face ea = sp.eaFace != null ? sp.eaFace : sp.face;
        return Math.max(sp.ascent + sp.descent, (ea.ascent + ea.descent) * 1.2) * t.size;
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
        t.bold = legacy;
        t.color = legacy ? Color.BLACK : new Color(0x595959);
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
            // 行の高さは文字の書体で決める (strRef の題名と同じ規則。exp46 の旧来のスタイル 18pt)。文字の枠の上の余白 (3.6pt) の分だけ上に描く
            TStyle tm = style(null);
            tm.size = (float) maxSize;
            tm.bold = t.bold;
            if (title != null) tm.apply(Xml.child(title, "txPr"));
            tm.size = (float) maxSize;
            String txt = richText(title);
            double h = lines * textTitleH(tm, txt);
            Rectangle2D tr = manualTitleRect(title, box, h);
            tr = new Rectangle2D.Double(tr.getX(), tr.getY() - RICH_TITLE_INSET, tr.getWidth(), tr.getHeight() + RICH_TITLE_INSET);
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
        // 題名の行の高さと基線は、文字に使う書体で決まる (Excel の印刷 exp72 の 14pt: 欧文だけの「V」は欧文の書体の
        // アセントで基線を置き、和文を含む「グラフ タイトル」は行が 0.464em 高く、基線が 0.06em 下)
        FontManager.Spec tsp = t.spec();
        boolean cjk = s.codePoints().anyMatch(FontManager::isCjk);
        double h = lineH(t) + (cjk ? TITLE_CJK_EXTRA * t.size : 0);
        titleExtra = cjk ? TITLE_CJK_EXTRA * t.size : 0;
        double by = box.getY() + (cjk ? TITLE_CJK_DROP * t.size : -Math.max(0, tsp.ascent - tsp.face.ascent) * t.size);
        // 行の高さを東アジアのフォントの寸法で広げた分 (メイリオなど) の一部だけ基線も下がる (Excel の印刷 exp92)
        by += (lineH(t) - (tsp.ascent + tsp.descent) * t.size) * 0.5;
        text(cv, s, box.getCenterX(), by, t, 0, -1);
        return h;
    }

    /** 表示単位のラベルの書式。 */
    private TStyle dispUnitStyle(Axis valAx) {
        TStyle tt = axisTitleStyle(valAx);
        Element lt = Xml.path(valAx.el, "dispUnits", "dispUnitsLbl", "txPr");
        if (lt != null) tt.apply(lt);
        return tt;
    }

    /** 表示単位のラベルの左の余白 (pt。Excel の印刷 exp72)。 */
    static final double DISP_UNIT_PAD = 4;

    /** 書式付きの題名の文字の枠の上の余白 (bodyPr の既定 tIns、pt)。 */
    static final double RICH_TITLE_INSET = 3.6;

    /** 題名の 1 行の高さ (和文を含めば lineH + 0.464em)。 */
    private double textTitleH(TStyle t, String s) {
        return lineH(t) + (s != null && s.codePoints().anyMatch(FontManager::isCjk) ? TITLE_CJK_EXTRA * t.size : 0);
    }

    /** 旧来のスタイルの値の軸のラベルの左に足す場所 (pt。exp75 の実測で 3pt 広い)。 */
    static final double LEGACY_LABEL_EXTRA = 3;
    /** 旧来のスタイルの第 2 軸のラベルは、さらに 5.8pt 右 (Excel の印刷 exp46)。 */
    static final double LEGACY_SEC_LABEL_EXTRA = 5.8;
    /** 系列が奥行きに並ぶ 3D の目盛りの最小の間隔 (通常の何倍か)。 */
    static final double RIBBON_VGAP = 1.2;
    /** 斜投影の 3D の系列名と床の右の縁の間 (pt。Excel の印刷 exp72)。 */
    static final double SER_LABEL_GAP = 10.4;
    /** 系列を奥に並べる 3D 縦棒 (系列 2〜3) の高さの倍率: 題名か凡例があれば 0.68 (Excel の印刷 exp72 の円柱・円錐・ピラミッド、exp50・52)、なければ 0.75 (exp73・74)。 */
    static final double RIBBON_TITLED = 0.68; // 箱の形で系列 3 なら 0.8 (exp88)
    /** 旧来のスタイルで軸がプロットの途中で交わるとき: 下の場所を 5.5pt 詰める、縦の軸が途中なら左に 9pt (Excel の印刷 exp47・48)。 */
    static final double LEGACY_CROSS_BOTTOM_CUT = 5.5, LEGACY_INSIDE_LEFT = 9;
    /** 旧来のスタイルのレーダーの値のラベルと上向きの線の間 (文字の大きさに対する比。Excel の印刷 exp75: 10pt で 9pt)。 */
    static final double LEGACY_RADAR_VAL_GAP = 0.9;
    /** 旧来のスタイルの四角の凡例の印: 大きさと左の空き (印の既定の大きさ 0.7em に対する比)。 */
    static final double LEGACY_SQ_KEY = 5.2 / 7, LEGACY_SQ_INSET = 3.0 / 7;
    /** 旧来のスタイルの縦の凡例の行の間隔 (文字の大きさに対する比) と、線の印の左の空き (pt)。Excel の印刷 exp75。 */
    static final double LEGACY_LEGEND_PITCH = 1.81, LEGACY_KEY_INSET = 2.6;
    /** 旧来のスタイルの縦の値のラベルは、さらに 4.5pt 左 (目盛の線の端から約 6.5pt。Excel の印刷 exp75)。 */
    static final double LEGACY_VAL_LABEL_OUT = 4.5;
    /** 上の凡例の線の印は 2.3pt 上 (Excel の印刷 exp81 のレーダー)。 */
    static final double LEGEND_TOP_RAISE = 2.3;
    /** 旧来のスタイル: 題名とプロットの間を 1pt 詰める、散布図の下を 1pt 多く空ける (Excel の印刷 exp75)。 */
    static final double LEGACY_TITLE_CUT = 1, LEGACY_XY_BOTTOM = 1;

    /** 旧来のスタイルの円の直径 ÷ プロットの短い辺 (exp75)。 */
    static final double LEGACY_PIE = 0.97, LEGACY_PIE_SHADE = 0.88;

    /** レーダーの項目名と外周の間 (pt。exp75)。 */
    static final double RADAR_LABEL_GAP = 2;
    /** 円・ドーナツの直径 = プロットの短い辺 − この値 (pt。Excel の印刷 exp81)。 */
    static final double PIE_INSET = 7;
    /** レーダーの値のラベルの右端と上向きの線の間 (文字の大きさに対する比。Excel の印刷 exp81: 9pt で 9.5pt)。 */
    static final double RADAR_VAL_GAP = 1.05;

    /** 線の系列のある凡例の印の幅 (文字の大きさに対する倍率)。 */
    static final double LEGEND_LINE_KEY = 2.3;
    /** 凡例の四角の印の大きさ (文字の大きさに対する比。Excel の印刷 exp81: 9pt で棒 5pt、円の要素 3pt)。 */
    static final double LEGEND_BAR_KEY = 5.0 / 9, LEGEND_PIE_KEY = 3.0 / 9;
    /** 下の凡例を下げる量と、四角の印を文字に対して下げる量 (pt、Excel の印刷 exp81)。 */
    static final double LEGEND_BOTTOM_DROP = 0.8, LEGEND_KEY_DROP = 0.7;
    /** 縦に並べる凡例の項目の間隔 (文字の大きさに対する比)。 */
    static final double LEGEND_V_PITCH = 1.875;
    /** 今の凡例の印が横長か。 */
    private boolean legendWide;

    /** 今のグラフの題名の和文の行の増分 (pt)。 */
    private double titleExtra;

    /** 和文を含む題名の行の高さの増分と基線の下げ (em)。 */
    static final double TITLE_CJK_EXTRA = 0.464, TITLE_CJK_DROP = 0.06;
    /** 項目ラベルの行の下に足す余白 (pt)。Excel の印刷 exp72: 9pt の欧文のラベルで 15.5pt = 行 11.56 + 4。 */
    static final double CAT_LABEL_PAD = 4;
    /** 上の軸 (横棒の第 2 軸) のラベルの行の下の余白 (pt)。 */
    static final double TOP_LABEL_PAD = 1.5;
    /** 値の軸が右 (crosses="max") のときのプロットの左の余白 (pt。exp72)。 */
    static final double CROSS_MAX_SHIFT = 3.6;

    /** タイトルの手動レイアウトの矩形。 */
    private Rectangle2D manualTitleRect(Element title, Rectangle2D box, double h) {
        return new Rectangle2D.Double(box.getX(), box.getY(), box.getWidth(), h + 2);
    }

    // ------------------------------------------------------------------ 凡例

    /** 複数行の横の凡例の列の数 (0 は 1 行) と列の幅。 */
    private int legCols;
    private double legColW;

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
        boolean pieLegend = false;
        List<String> names = new ArrayList<>();
        List<Color> colors = new ArrayList<>();
        List<Boolean> lineSym = new ArrayList<>();
        List<Series> legSeries = new ArrayList<>();
        List<Boolean> markSym = new ArrayList<>(); // 線の印にマーカーも重ねる系列 (マーカー付きの折れ線。Excel の印刷 exp46)
        java.util.Set<Integer> hollow = new java.util.HashSet<>();
        java.util.Set<Integer> deleted = new java.util.HashSet<>();
        for (Element le : Xml.children(legend, "legendEntry")) {
            if (Xml.boolAttr(Xml.child(le, "delete"), "val", false)) deleted.add(Xml.intAttr(Xml.child(le, "idx"), "val", -1));
        }
        Group g0 = groups.get(0);
        if (g0.type.equals("surface") && !g0.series.isEmpty()) {
            // 等高線: 値の目盛りの区間ごとの色 (Excel の凡例は「0-20」「20-40」…)。目盛りはプロットの高さの概算で先に決める
            Axis va = null;
            for (Axis a : chartAxes) if (a.kind.equals("valAx")) va = a;
            if (va == null) { va = new Axis(); va.kind = "valAx"; }
            va.noHeadroom = true;
            double mn = Double.POSITIVE_INFINITY, mx = Double.NEGATIVE_INFINITY;
            for (Series s : g0.series) for (Double v : s.vals) if (v != null) { mn = Math.min(mn, v); mx = Math.max(mx, v); }
            if (topViewChart()) autoScale(va, mn, mx, TOP_VIEW_BANDS, 1);
            else autoScale(va, mn, mx, plot.getHeight() * 0.5, 9 * PERSP_VGAP);
            String fmt = g0.series.get(0).fmt == null ? "General" : g0.series.get(0).fmt;
            int k = 0;
            for (double v = va.lo; v < va.hi - va.unit * 1e-6 && k < 40; v = step(va, v), k++) {
                names.add(axisLabel(va, v, fmt) + "-" + axisLabel(va, step(va, v), fmt));
                colors.add(autoColor(k)); // ワイヤーフレームも元の色の枠 (exp72)
                if (g0.wireframe) hollow.add(names.size() - 1);
                lineSym.add(false);
            }
        } else if ((g0.type.equals("pie") || g0.type.equals("doughnut")) || (g0.varyColors && g0.series.size() == 1 && !g0.series.isEmpty())) {
            pieLegend = true;
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
                        while (legSeries.size() < names.size() - 1) legSeries.add(null);
                        legSeries.add(s);
                        lineSym.add(g.type.equals("line") && !g.is3D || (g.type.equals("scatter") && !g.scatterStyle.equals("marker")) || g.type.equals("radar")); // 3D の折れ線は四角の印 (exp88)
                        while (markSym.size() < names.size() - 1) markSym.add(false);
                        markSym.add((g.type.equals("line") && !g.is3D || g.type.equals("scatter") && !g.scatterStyle.equals("marker")) && g.markers);
                    }
                    k++;
                }
            }
            // 横棒の縦の凡例は、棒の並び (下から系列順) に合わせて逆順 (Excel の印刷 exp75)
            boolean allHoriz = true;
            for (Group g : groups) allHoriz &= g.type.equals("bar") && g.horizontal;
            if (allHoriz && (pos.equals("r") || pos.equals("l") || pos.equals("tr"))) {
                java.util.Collections.reverse(names);
                java.util.Collections.reverse(colors);
                java.util.Collections.reverse(lineSym);
                while (markSym.size() < names.size()) markSym.add(false);
                java.util.Collections.reverse(markSym);
                while (legSeries.size() < names.size()) legSeries.add(null);
                java.util.Collections.reverse(legSeries);
            }
        }
        if (names.isEmpty()) return plot;
        double lh = lineH(t) * 1.05;
        double sym = t.size * 0.7;
        // 線の系列があれば、すべての印を文字の大きさ × 2.3 の幅にし (棒は横長の四角)、印と文字の間を 1.5pt、項目の間を 7pt にする (Excel の印刷 exp72・exp75)
        boolean wideKeys = false;
        for (int i = 0; i < lineSym.size(); i++)
            if (lineSym.get(i) && !(i < legSeries.size() && legSeries.get(i) != null && lineNone(legSeries.get(i)))) wideKeys = true;
        double symW = wideKeys ? t.size * LEGEND_LINE_KEY : sym, keyGap = wideKeys ? 1.5 : 4, itemGap = wideKeys ? 7 : 10;
        boolean squareKeys = false;
        if (!wideKeys && !legacy) { // 等高線の区間の凡例も同じ四角の印 (exp72)
            // 四角の印 (Excel の印刷 exp81、9pt): 棒などは 5pt、印と文字の間 2.2pt、項目の間 4.9pt。円・ドーナツの要素は 3pt、3pt、7.1pt
            sym = symW = t.size * (pieLegend && !dp.bw ? LEGEND_PIE_KEY : LEGEND_BAR_KEY); // 白黒印刷は円の要素も 5/9em
            keyGap = pieLegend ? 3.0 : 2.2;
            itemGap = pieLegend ? 7.1 : 4.9;
            squareKeys = true;
        }
        legendWide = wideKeys;
        double maxW = 0;
        for (String n : names) maxW = Math.max(maxW, textWidth(n, t));
        boolean vertical = pos.equals("r") || pos.equals("l") || pos.equals("tr");
        Element ml = Xml.path(legend, "layout", "manualLayout");
        Rectangle2D box;
        boolean modernWide = wideKeys && !legacy;
        if (vertical && (squareKeys || modernWide)) lh = t.size * LEGEND_V_PITCH; // 縦に並べる四角の印・線の印の間隔 (Excel の印刷 exp81: 9pt で 16.9pt、exp89)
        if (vertical && legacy) lh = t.size * LEGACY_LEGEND_PITCH; // 旧来のスタイル: 10pt で 18.1pt (exp75)
        if (vertical) {
            double w = symW + keyGap + maxW + 6;
            double h = names.size() * lh + 4;
            double x = pos.equals("l") ? inner.getX() : inner.getMaxX() - w;
            double y = plot.getY() + (plot.getHeight() - h) / 2 - (squareKeys || modernWide ? 1.1 : 0);
            box = new Rectangle2D.Double(x, y, w, h);
            if (!overlay) {
                double ex = squareKeys || modernWide ? 3 : 0; // 四角の印・線の印の凡例はプロットとの間を 3pt 広く (Excel の印刷 exp81・89)
                if (pos.equals("l")) plot = new Rectangle2D.Double(plot.getX() + w + 6 + ex, plot.getY(), plot.getWidth() - w - 6 - ex, plot.getHeight());
                else plot = new Rectangle2D.Double(plot.getX(), plot.getY(), plot.getWidth() - w - 6 - ex, plot.getHeight());
            }
        } else {
            double total = 0;
            for (String n : names) total += symW + keyGap + textWidth(n, t) + itemGap;
            // 最後の項目の後ろは今までどおり。四角の印 (squareKeys) は印から最後の文字までを中央に置く (Excel の印刷 exp81)
            // 線の印 (wideKeys) も同じ (exp81 の折れ線・レーダー)
            total -= squareKeys || wideKeys && !legacy ? itemGap - 3.4 : itemGap - 10;
            // 1 行に収まらなければ、最も長い項目の幅の列に並べて複数行にする (Excel の印刷 exp125 dleg・pleg)
            legCols = 0;
            double avail = inner.getWidth();
            if (total > avail && names.size() > 1 && ml == null) {
                // 行数は 1 行の長さ ÷ 幅 (切り上げ) から始め、その行数で列が幅に収まるまで増やす (en2 の 180・230・300、cjk4_180)
                legColW = symW + keyGap + maxW + itemGap;
                int r = (int) Math.ceil(total / avail);
                legCols = (names.size() + r - 1) / r;
                while (legCols > 1 && legCols * legColW - itemGap > avail) { r++; legCols = (names.size() + r - 1) / r; }
                total = legCols * legColW - itemGap + (squareKeys || wideKeys && !legacy ? 3.4 : 10);
            }
            int rows = legCols > 0 ? (names.size() + legCols - 1) / legCols : 1;
            double h = lh * rows + 4;
            double y = pos.equals("t") ? plot.getY() - (wideKeys && !legacy ? LEGEND_TOP_RAISE : 0)
                    : inner.getMaxY() - h + (squareKeys || wideKeys ? LEGEND_BOTTOM_DROP : 0);
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
        dumpLegend = box;
        Element spPr = Xml.child(legend, "spPr");
        if (spPr != null) {
            dp.fillShape(cv, box, dp.paint(DrawingPainter.fillElement(spPr), null, null), null);
            dp.strokeShape(cv, box, dp.line(Xml.child(spPr, "ln"), null, null));
        }
        double x = box.getX() + 3, y = box.getY() + 2;
        for (int i = 0; i < names.size(); i++) {
            if (vertical) {
                if (hollow.contains(i)) cv.stroke(new Rectangle2D.Double(x + (symW - sym) / 2, y + lh / 2 - sym / 2, sym, sym), new PdfCanvas.Stroke(0.75f), colors.get(i));
                else legendSymbol(cv, x, y + lh / 2 + (squareKeys ? LEGEND_KEY_DROP : 0), symW, sym, colors.get(i), lineSym.get(i), i < legSeries.size() ? legSeries.get(i) : null);
                if (i < markSym.size() && markSym.get(i) && i < legSeries.size() && legSeries.get(i) != null && !lineNone(legSeries.get(i)))
                    marker(cv, legSeries.get(i), i, x + (legacy ? (LEGACY_KEY_INSET + symW) / 2 : symW / 2), y + lh / 2);
                text(cv, names.get(i), x + symW + keyGap, y + lh / 2, t, -1, 0);
                y += lh;
            } else {
                if (hollow.contains(i)) cv.stroke(new Rectangle2D.Double(x + (symW - sym) / 2, y + lh / 2 - sym / 2, sym, sym), new PdfCanvas.Stroke(0.75f), colors.get(i));
                else legendSymbol(cv, x, y + lh / 2 + (squareKeys ? LEGEND_KEY_DROP : 0), symW, sym, colors.get(i), lineSym.get(i), i < legSeries.size() ? legSeries.get(i) : null);
                if (i < markSym.size() && markSym.get(i) && i < legSeries.size() && legSeries.get(i) != null && !lineNone(legSeries.get(i)))
                    marker(cv, legSeries.get(i), i, x + symW / 2, y + lh / 2);
                text(cv, names.get(i), x + symW + keyGap, y + lh / 2, t, -1, 0);
                x += symW + keyGap + textWidth(names.get(i), t) + itemGap;
                if (legCols > 0) {
                    x = box.getX() + 3 + ((i + 1) % legCols) * legColW;
                    if ((i + 1) % legCols == 0) y += lh;
                }
            }
        }
        return plot;
    }

    /**
     * 凡例の記号。線の系列で線が「なし」なら、マーカーだけを描く (マーカーもなければ描かない。Excel の印刷 exp72 の株価)。
     */
    private void legendSymbol(PdfCanvas cv, double x, double y, double w, double sym, Color c, boolean line, Series s) throws IOException {
        if (line && s != null && lineNone(s)) {
            marker(cv, s, 0, x + w / 2, y);
            return;
        }
        drawSymbol(cv, x, y, w, sym, c, line);
    }

    /** 凡例の記号 (線またはマーカー、四角) を描く。 */
    private void drawSymbol(PdfCanvas cv, double x, double cy, double w, double h, Color c, boolean line) throws IOException {
        if (line && legacy && legendWide) {
            // 旧来のスタイルの線の印: 2.75pt の太さで、2.6pt 右から始まる (Excel の印刷 exp75)
            cv.line(x + LEGACY_KEY_INSET, cy, x + w, cy, new PdfCanvas.Stroke(2.75f), c);
        } else if (line) {
            cv.line(x, cy, x + (legendWide ? w : w * 1.6), cy, new PdfCanvas.Stroke(2f), c);
        } else if (legendWide) {
            cv.fillRect(x, cy - h * 0.35, w, h * 0.7, c);
        } else if (legacy && !dp.bw) {
            // 旧来のスタイルの四角の印: 0.52em 角で、0.3em 右 (10pt で 5.2pt 角、3pt。Excel の印刷 exp75)
            double k = h * LEGACY_SQ_KEY;
            cv.fillRect(x + h * LEGACY_SQ_INSET, cy - k / 2, k, k, c);
        } else {
            cv.fillRect(x, cy - h / 2, w, h, c);
            // 白黒印刷: 四角の印に黒の枠 (Excel の印刷 exp59 の bw)
            if (dp.bw) cv.stroke(new Rectangle2D.Double(x, cy - h / 2, w, h), new PdfCanvas.Stroke(0.5f), Color.BLACK);
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

    /** 自動の目盛りの区間の数の上限 (Excel の印刷 exp72 のバブル: −1〜10 の 1 刻み、−4〜18 の 2 刻み)。 */
    static final int MAX_INTERVALS = 11;

    /** Excel に似た値の軸の自動目盛り。 */
    static void autoScale(Axis a, double dmin, double dmax, double lengthPt) {
        autoScale(a, dmin, dmax, lengthPt, 28);
    }

    /**
     * Excel に似た値の軸の自動目盛り。
     *
     * @param gap 目盛りの間隔の最小 (pt)。縦の軸は目盛ラベルの高さで決まる (Excel の印刷: 10pt の文字で 0〜7 を 1 刻み)
     */
    static void autoScale(Axis a, double dmin, double dmax, double lengthPt, double gap) {
        if (Double.isInfinite(dmin)) { dmin = 0; dmax = 1; }
        if (a.logScale) {
            // 対数目盛: 底のべき乗で区切る (正の値だけを使う)
            double b = a.logBase;
            double pmin = dmin > 0 ? dmin : 1, pmax = Math.max(dmax, pmin);
            a.lo = a.min != null && a.min > 0 ? a.min : Math.pow(b, Math.floor(Math.log(pmin) / Math.log(b) + 1e-9));
            a.hi = a.max != null && a.max > a.lo ? a.max : Math.pow(b, Math.ceil(Math.log(pmax) / Math.log(b) - 1e-9));
            if (a.hi <= a.lo) a.hi = a.lo * b;
            a.unit = b;
            return;
        }
        double lo = a.min != null ? a.min : Math.min(0, dmin);
        double hi = a.max != null ? a.max : Math.max(0, dmax);
        if (a.min == null && dmin > 0 && (dmax - dmin) / dmax < 1.0 / 6) lo = dmin - (dmax - dmin) / 2; // Excel は最小値を引き上げる
        if (hi == lo) hi = lo + 1;
        // 目盛りの間隔: 1, 2, 5 × 10^n の小さい方から試し、丸めた範囲で目盛りの間隔が gap 以上になる最初のもの
        // (Excel の印刷: 0〜45 を 5 刻み、-4〜8 を 2 刻み。目盛りは 10 区間まで)
        double unit = a.majorUnit != null ? a.majorUnit : niceUnit(hi - lo, 10);
        double lo0 = lo, hi0 = hi;
        for (int iter = 0; iter < 40; iter++) {
            lo = lo0;
            hi = hi0;
            if (a.min == null) lo = Math.floor(lo / unit) * unit;
            if (a.max == null) {
                double h2 = Math.ceil(hi / unit) * unit;
                if (!a.noHeadroom && dmax > 0 && h2 - dmax < (a.exactHeadroom ? 1e-9 * Math.max(1, Math.abs(h2)) : 0.05 * (h2 - lo))) h2 += unit; // 上端の余裕
                hi = h2;
            }
            if (a.min == null && dmin < 0) {
                double l2 = Math.floor(dmin / unit) * unit;
                if (!a.noHeadroom && dmin < 0 && dmin - l2 < 0.05 * (hi - l2)) l2 -= unit;
                lo = l2;
            }
            double intervals = (hi - lo) / unit;
            if (a.majorUnit != null || (intervals <= a.maxIntervals + 0.001 && lengthPt / intervals >= gap) || intervals <= 2.001) break;
            unit = nextNice(unit);
        }
        a.lo = lo;
        a.hi = hi;
        a.unit = unit;
    }

    /** 1, 2, 5 × 10^n の次の値。 */
    static double nextNice(double u) {
        double mag = Math.pow(10, Math.floor(Math.log10(u) + 1e-9));
        double n = u / mag;
        return n < 1.5 ? 2 * mag : n < 3.5 ? 5 * mag : 10 * mag;
    }

    /** 軸の目盛ラベルの文字列。 */
    private String axisLabel(Axis a, double v, String srcFmt) {
        String fmt = a.fmt != null ? a.fmt : srcFmt;
        v /= a.dispUnit;
        if (Math.abs(v) < 1e-12) v = 0;
        return ctx.formatter.formatNumber(v, fmt == null ? "General" : fmt);
    }

    /**
     * 第 2 軸 (値の軸が右か上) のグループが先に保存されているとき、主軸のグループを先頭にする。
     * 例: 折れ線 (第 2 軸) + 散布図 (主軸) の順の組み合わせ (rand30 rd09。先頭で軸を決めると散布図の X 軸がなかった)。
     */
    private static List<Group> primaryFirst(List<Group> groups, List<Axis> axes) {
        if (groups.size() < 2 || !isSecondary(groups.get(0), axes)) return groups;
        for (int i = 1; i < groups.size(); i++) {
            if (isSecondary(groups.get(i), axes)) continue;
            List<Group> r = new ArrayList<>(groups);
            r.add(0, r.remove(i));
            return r;
        }
        return groups;
    }

    /** グループの値の軸 (散布図は Y 軸) が右か上にある。 */
    private static boolean isSecondary(Group gg, List<Axis> axes) {
        boolean sc = gg.type.equals("scatter");
        for (Axis a : axes) {
            if (!gg.axIds.contains(a.id) || !a.kind.equals("valAx")) continue;
            if (sc && (a.pos.equals("b") || a.pos.equals("t"))) continue;
            return a.pos.equals("r") || a.pos.equals("t");
        }
        return false;
    }

    /** 直交座標のグラフ (棒、折れ線、面、散布図) を描く: 目盛り、目盛線、系列、軸、ラベル。 */
    private void paintCartesian(PdfCanvas cv, List<Group> groups, List<Axis> axes, Rectangle2D box, boolean innerLayout,
                                Element plotSpPr) throws IOException {
        groups = primaryFirst(groups, axes);
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

        midCat = !scatter && "midCat".equals(valAx.crossBetween);
        flat3D = g.is3D;
        valAx.noHeadroom = g.is3D || g.type.equals("surface");
        // 第 2 軸: 最初のグループと違う値の軸を使うグループ (散布図・横棒以外の組み合わせも同じ)
        List<Group> primary = new ArrayList<>(), secondary = new ArrayList<>();
        Axis valAx2 = null;
        for (Group gg : groups) {
            Axis own = null;
            if (gg != g && !gg.axIds.isEmpty() && !gg.axIds.contains(valAx.id)) {
                for (Axis a : axes) if (gg.axIds.contains(a.id) && a.kind.equals("valAx")) own = a;
            }
            if (own == null) primary.add(gg);
            else { secondary.add(gg); if (valAx2 == null) valAx2 = own; }
        }

        // データの範囲
        int nCats = 0;
        for (Group gg : groups) for (Series s : gg.series) nCats = Math.max(nCats, Math.max(s.vals.size(), s.cats.size()));
        double[] range = dataRange(primary, nCats);
        double dmin = range[0], dmax = range[1];
        // 近似曲線の値も値の軸の範囲に含める (Excel と同じ)
        for (Group gg : primary) for (Series s : gg.series) {
            double[] tr = trendRange(s);
            if (tr != null) { dmin = Math.min(dmin, tr[0]); dmax = Math.max(dmax, tr[1]); }
            // 誤差範囲の端も含める (Excel の印刷: 14 の +20% → 最大 18)
            if (s.ser != null) for (Element eb : Xml.children(s.ser, "errBars")) {
                if ("x".equals(Xml.attr(Xml.child(eb, "errDir"), "val", "y"))) continue;
                for (double[] lh : errorSpans(s, eb)) if (lh != null) { dmin = Math.min(dmin, lh[0]); dmax = Math.max(dmax, lh[1]); }
            }
        }
        // バブル: 最大のバブルの半径 ÷ プロットの短い辺 = R の分だけ、x・y の両方の範囲を広げる (範囲 ÷ (1 − 2R)。Excel の印刷 exp72:
        // 倍率 25〜300% の x・y の目盛りの範囲と一致)
        double bubR = g.bubble ? Math.min(0.4, BUBBLE_A * Math.max(0, g.bubbleScale) / (Math.max(0, g.bubbleScale) + BUBBLE_K) / 2) : 0;
        if (bubR > 0 && dmax > dmin) {
            double span = (dmax - dmin) / (1 - 2 * bubR);
            dmin -= bubR * span; dmax += bubR * span;
        }
        if (bubR > 0) valAx.noHeadroom = true;
        String srcFmt = g.series.isEmpty() ? "General" : g.series.get(0).fmt;
        if (g.grouping.equals("percentStacked")) { srcFmt = "0%"; valAx.noHeadroom = true; }
        double[] range2 = valAx2 == null ? null : dataRange(secondary, nCats);
        String srcFmt2 = secondary.isEmpty() || secondary.get(0).series.isEmpty() ? "General" : secondary.get(0).series.get(0).fmt;
        if (!secondary.isEmpty() && secondary.get(0).grouping.equals("percentStacked")) srcFmt2 = "0%";
        TStyle vt2 = valAx2 == null ? null : style(Xml.child(valAx2.el, "txPr"));
        boolean val2Visible = valAx2 != null && !valAx2.deleted && !"none".equals(valAx2.tickLblPos);

        TStyle vt = style(Xml.child(valAx.el, "txPr"));
        TStyle ct = catAx != null ? style(Xml.child(catAx.el, "txPr")) : vt;
        if (scatter) ct = style(Xml.child(xValAx.el, "txPr"));

        // ラベルの幅を測るための仮の目盛り
        double valLen = horiz ? box.getWidth() : box.getHeight();
        // 目盛りの間隔の最小: 縦の軸は目盛ラベルの高さ、横の軸は 36pt (Excel の印刷の実測)
        // (書式のないグラフ (旧来のスタイル) は文字の大きさ × 1.25。日本語のフォントの行の高さは大きく、Excel より粗くなるため)
        double vgap = horiz ? hGap() : legacy ? vt.size * LEGACY_GAP : lineH(vt) * 0.85;
        autoScale(valAx, dmin, dmax, valLen, vgap);
        double xmin = Double.POSITIVE_INFINITY, xmax = Double.NEGATIVE_INFINITY;
        if (scatter) {
            for (Series s : g.series) {
                for (int i = 0; i < s.vals.size(); i++) {
                    double x = i < s.xs.size() && s.xs.get(i) != null ? s.xs.get(i) : i + 1;
                    xmin = Math.min(xmin, x);
                    xmax = Math.max(xmax, x);
                }
            }
            if (bubR > 0 && xmax > xmin) {
                double span = (xmax - xmin) / (1 - 2 * bubR);
                xmin -= bubR * span; xmax += bubR * span;
            }
            if (bubR > 0) xValAx.noHeadroom = true;
            // 書式のあるグラフは、y のラベルの分 (20pt) を除いた長さで目盛りを決める (Excel の印刷 exp71・72・81)
            autoScale(xValAx, xmin, xmax, box.getWidth() - (!modern() ? 0 : H_LEN_CUT), hGap());
        }

        // 回転と遠近で投影する 3D (直交軸でない縦棒・横棒・折れ線・面・等高線)
        if (g.type.equals("surface") || g.is3D && !rAngAx && !scatter && (g.type.equals("bar") || g.type.equals("line") || g.type.equals("area"))) {
            Axis serAx = null;
            for (Axis a : axes) if (a.kind.equals("serAx") && (g.axIds.isEmpty() || g.axIds.contains(a.id))) serAx = a;
            paintPersp(cv, g, catAx, valAx, serAx, box, dmin, dmax, srcFmt, vt, ct);
            return;
        }

        // ラベルの場所を空ける
        Rectangle2D plot = box;
        if (valAx2 != null) autoScale(valAx2, range2[0], range2[1], valLen, vgap);
        if (!innerLayout) {
            double left = 0, bottom = 0, right = 0, topSpace = 0;
            if (val2Visible) {
                if (!horiz) {
                    double mw = 0;
                    for (double v = valAx2.lo; v <= valAx2.hi + valAx2.unit * 1e-6; v = step(valAx2, v)) mw = Math.max(mw, textWidth(axisLabel(valAx2, v, srcFmt2), vt2));
                    // 第 2 軸のラベル: プロットとの間は 0.91em + 0.8pt、右の場所はそれ − 3.1pt (Excel の印刷 exp72)
                    right = modern() ? mw + valLabelGap(vt2) + SEC_LABEL_EXTRA - SEC_RIGHT_CUT : mw + 6;
                } else topSpace = textLineH(vt2, List.of("0")) + TOP_LABEL_PAD - (modern() ? 1.2 : 0); // exp72 の sec_bar
            }
            if (valAx2 != null && valAx2.titleEl != null) {
                if (!horiz) right += lineH(vt2) + 4; else topSpace += lineH(vt2) + 4;
            }
            boolean valVisible = !valAx.deleted && !"none".equals(valAx.tickLblPos);
            boolean catVisible = scatter ? !xValAx.deleted : catAx != null && !catAx.deleted && !"none".equals(catAx.tickLblPos);
            if (!horiz) {
                if (valVisible) {
                    double mw = 0;
                    for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v = step(valAx, v)) mw = Math.max(mw, textWidth(axisLabel(valAx, v, srcFmt), vt));
                    // 交差位置が最大 (crosses="max") なら、値の軸と目盛ラベルはプロットの右
                    // 値の軸がプロットの途中の項目で交わるときは、ラベルは軸の横に描くので左に場所を取らない (Excel の印刷 exp72)
                    boolean nextTo = "nextTo".equals(valAx.tickLblPos == null ? "nextTo" : valAx.tickLblPos);
                    boolean inside = !scatter && valAx.crossesAt != null && valAx.crossesAt > 1 && nextTo
                            || scatterZeroInside(scatter, valAx, xValAx) && nextTo;
                    if ("max".equals(valAx.crosses) && !scatter) {
                        right += (modern() ? mw + valLabelGap(vt) - VAL_LABEL_OUT : mw + 6) - CROSS_MAX_SHIFT; left = CROSS_MAX_SHIFT;
                    } else if (inside && modern()) left = INSIDE_LEFT; // 軸が途中で交わるときも左に 3.5pt (Excel の印刷 exp72)
                    else if (inside && legacy) left = LEGACY_INSIDE_LEFT;
                    else if (!inside) left = mw + valLabelGap(vt) + (!modern() ? 2 : -VAL_LABEL_OUT) + (legacy ? LEGACY_LABEL_EXTRA : 0);
                }
                if (catVisible) bottom = textLineH(ct, List.of("0")) + CAT_LABEL_PAD + (scatter && modern() ? 1.3 : 0);
                // 散布図の x の軸が y = 0 で交わる (y の範囲が 0 をまたぐ) ときは、ラベルを軸の横に描くので下に場所を取らない
                if (catVisible && scatterZeroInside(scatter, xValAx, valAx))
                    bottom = !valVisible ? 4 : modern() ? vt.size * NEG_BOTTOM : lineH(vt) / 2 + 1 - LEGACY_CROSS_BOTTOM_CUT; // 項目軸の交差と同じ (exp72 のバブル、exp48)
                if (catVisible && !scatter) bottom = catLabelHeight(g, catAx, ct, (box.getWidth() - left) / Math.max(1, nCats)) + CAT_LABEL_PAD
                        + catBottomExtra(ct, g.series.isEmpty() ? List.of() : g.series.get(0).cats);
                // 項目軸がプロットの途中で交わる (負の値があり 0 で交わる) とき、ラベルは軸の横に描くので、下に場所を取らない
                // (Excel の印刷。下には一番下の値のラベルの半分だけ)
                if (catVisible && !scatter && catAx != null && !"low".equals(catAx.tickLblPos)) {
                    double crossV = "max".equals(catAx.crosses) ? valAx.hi : "min".equals(catAx.crosses) ? valAx.lo
                            : catAx.crossesAt != null ? catAx.crossesAt : 0;
                    if (crossV > valAx.lo + 1e-9 && !valAx.reverse)
                        bottom = valVisible ? (modern() ? vt.size * NEG_BOTTOM : lineH(vt) / 2 + 1 - LEGACY_CROSS_BOTTOM_CUT) : 4; // 書式のあるグラフは 0.53em (exp81 の負の値)
                }
            } else {
                if (catVisible) {
                    double mw = 0;
                    Series s0 = g.series.isEmpty() ? null : g.series.get(0);
                    if (s0 != null) for (String c : s0.cats) mw = Math.max(mw, textWidth(c, ct));
                    left = mw + valLabelGap(ct) + (!modern() ? 2 : -VAL_LABEL_OUT) + (legacy ? LEGACY_LABEL_EXTRA : 0);
                }
                if (valVisible) bottom = textLineH(vt, List.of("0")) + CAT_LABEL_PAD;
            }
            // 軸のタイトル: 縦の軸 (左) は回して左に、横の軸 (下) は下に場所を取る
            Axis leftAx = horiz ? catAx : valAx, bottomAx = scatter ? xValAx : horiz ? valAx : catAx;
            if (leftAx != null && leftAx.titleEl != null)
                left += horizontalTitle(leftAx) ? textWidth(axisTitleText(leftAx), titleSize(axisTitleStyle(leftAx))) + H_TITLE_GAP
                        : lineH(horiz ? ct : vt) + AXIS_TITLE_GAP;
            else if (!horiz && valAx.dispUnitLbl && valAx.dispUnitName != null)
                left += lineH(horiz ? ct : vt) + 4 + TITLE_CJK_EXTRA * dispUnitStyle(valAx).size + DISP_UNIT_PAD; // 和文の行 (exp72 の実測)
            if (bottomAx != null && bottomAx.titleEl != null && !axisTitleText(bottomAx).isEmpty())
                bottom += lineH(titleSize(axisTitleStyle(bottomAx))) + AXIS_TITLE_GAP;
            if (dTable != null && !horiz && !scatter) {
                // データテーブル: 項目名の行と系列の行をプロットの下に置く。左の列 (凡例の印と系列名) はプロットの左へはみ出す
                double rowH = dtRowH(groups, Xml.child(dTable, "txPr") == null ? ct : style(Xml.child(dTable, "txPr")));
                bottom = rowH * (1 + allSeries(groups).size()) - DT_BOTTOM;
                left = Math.max(left, dtLeftWidth(groups, ct) + DT_INSET);
            }
            // 横の値の軸 (横棒・散布図の x・横棒の上の第 2 軸) は、最後のラベルの半分 + 3pt を右に空ける (Excel の印刷 exp72)
            double rOver = 0;
            Axis hv = scatter ? xValAx : horiz ? valAx : null;
            if (hv != null && !hv.deleted && !"none".equals(hv.tickLblPos))
                rOver = Math.max(rOver, textWidth(axisLabel(hv, hv.hi, scatter ? "General" : srcFmt), scatter ? ct : vt) / 2 + 3);
            if (horiz && val2Visible) rOver = Math.max(rOver, textWidth(axisLabel(valAx2, valAx2.hi, srcFmt2), vt2) / 2 + 3);
            // 項目を目盛りの上に置く (midCat) 横の項目軸も、最後の項目名の半分 + 3pt (Excel の印刷 exp75 の面)
            if (midCat && !horiz && catVisible && !g.series.isEmpty() && !g.series.get(0).cats.isEmpty()) {
                List<String> cs = g.series.get(0).cats;
                rOver = Math.max(rOver, textWidth(cs.get(cs.size() - 1), ct) / 2 + 3);
            }
            right = Math.max(right, rOver - 4);
            // 最初のラベルも左に半分 + 3pt (y のラベルが軸の横に入り、左に場所を取らないとき。exp72 のバブル)
            if (hv != null && !hv.deleted && !"none".equals(hv.tickLblPos))
                left = Math.max(left, textWidth(axisLabel(hv, hv.lo, scatter ? "General" : srcFmt), scatter ? ct : vt) / 2 + 3);
            if (legacy && scatter) bottom += LEGACY_XY_BOTTOM; // 旧来のスタイルの散布図は x のラベルの下に 1pt 多く空く (exp75)
            double htop = horiz && modern() ? H_PLOT_TOP : 0; // 書式のある横棒はプロットの上端が 0.6pt 上 (Excel の印刷 exp89)
            plot = new Rectangle2D.Double(box.getX() + left, box.getY() + 4 + topSpace - htop, Math.max(10, box.getWidth() - left - 4 - right),
                    Math.max(10, box.getHeight() - bottom - 4 - topSpace + htop));
            autoScale(valAx, dmin, dmax, horiz ? plot.getWidth() : plot.getHeight(), vgap);
            if (valAx2 != null) autoScale(valAx2, range2[0], range2[1], horiz ? plot.getWidth() : plot.getHeight(), vgap);
        }

        // プロットエリアの塗りつぶし
        if (plotSpPr != null) {
            dp.fillShape(cv, plot, dp.paint(DrawingPainter.fillElement(plotSpPr), null, null), null);
        }
        // 値の軸の位置 (縦の軸): 値の軸の crosses / crossesAt (項目軸の上の位置。項目の番号) で決まる
        valAxisX = plot.getX();
        if (scatterZeroInside(scatter, valAx, xValAx)) valAxisX = plot.getX() + (0 - xValAx.lo) / (xValAx.hi - xValAx.lo) * plot.getWidth();
        if (!horiz && !scatter && catAx != null) {
            if ("max".equals(valAx.crosses)) valAxisX = plot.getMaxX();
            else if (valAx.crossesAt != null && nCats > 0) {
                double k = Math.max(1, Math.min(nCats, valAx.crossesAt)) - 1;
                valAxisX = midCat ? plot.getX() + k * plot.getWidth() / Math.max(1, nCats - 1) : plot.getX() + k * plot.getWidth() / nCats;
            }
        }
        // 3D の縦棒: 奥の壁の分だけ右と上を空け、奥行きの向き (斜投影。Excel の直交軸の既定) を決める
        d3x = 0; d3y = 0;
        boolean ribbon = g.is3D && ((g.type.equals("line") || g.type.equals("area") || g.type.equals("bar")) && "standard".equals(g.grouping)
                || g.type.equals("surface"));
        if (g.is3D && (g.type.equals("bar") && !(horiz && ribbon) || ribbon && !horiz) && nCats > 0) {
            // 直交軸の 3D (斜投影。Excel の印刷 exp73): 奥行き D の向きは (D sin rotY, −D sin rotX)。D は項目の幅 × m × 奥行き%
            // (集合: (1 + gapDepth) / (系列数 + gapWidth)、積み上げ: (1 + gapDepth) / (1 + gapWidth)、系列を奥に並べる: 系列数)。
            // 正面の高さ (横棒は項目の向きの長さ) = 幅 × 比 (高さ% の指定、なければプロットの縦 ÷ 横 × 倍率)。
            // 正面と奥行きの全体がプロットに収まる最大の大きさにして、中央に置く
            if (!innerLayout) {
                // 外側の余白 (グラフの内側の枠から。ラベルの場所と大きい方。exp73 の実測)
                // ラベルがあれば、左に 4pt、下に 3.5pt 多く取る。横棒の値のラベルは右に 13pt、奥の系列のラベルは右に系列名 + 8.5pt (exp74)
                double l0 = plot.getX() - box.getX(), b0 = box.getMaxY() - plot.getMaxY();
                double l = Math.max(l0 > 0 ? l0 + OBLIQUE_LABEL_LEFT : 0, OBLIQUE_INSET[0]), rt = Math.max(box.getMaxX() - plot.getMaxX(), OBLIQUE_INSET[1]);
                double tp = Math.max(plot.getY() - box.getY(), OBLIQUE_INSET[2]), bt = Math.max(b0 > 4.01 ? b0 + OBLIQUE_LABEL_BOTTOM : 0, OBLIQUE_INSET[3]);
                if (horiz && !valAx.deleted && !"none".equals(valAx.tickLblPos)) rt = Math.max(rt, OBLIQUE_BAR_RIGHT);
                if (ribbon) for (Axis a : axes) if (a.kind.equals("serAx") && (g.axIds.isEmpty() || g.axIds.contains(a.id)) && !a.deleted && !"none".equals(a.tickLblPos)) {
                    TStyle st = style(Xml.child(a.el, "txPr"));
                    double sw = 0;
                    for (Series ss : g.series) sw = Math.max(sw, textWidth(ss.name, st));
                    rt = Math.max(rt, sw + OBLIQUE_INSET[1]);
                }
                plot = new Rectangle2D.Double(box.getX() + l, box.getY() + tp, Math.max(10, box.getWidth() - l - rt), Math.max(10, box.getHeight() - tp - bt));
            }
            int ns = Math.max(1, g.series.size());
            double gd = g.gapDepth / 100, gw = g.gapWidth / 100;
            double m = (ribbon ? ns : "clustered".equals(g.grouping) ? (1 + gd) / (ns + gw) : (1 + gd) / (1 + gw)) * depthPct / 100;
            double sy = Math.sin(Math.toRadians(rotY)), sx = Math.sin(Math.toRadians(rotX));
            double factor = !ribbon ? 1 : g.type.equals("bar") ? (ns <= 1 ? 0.5 : ns < 4 ? (Xml.child(chartEl, "legend") != null || Xml.child(chartEl, "title") != null ? ("box".equals(g.shape) && ns == 3 ? 0.8 : RIBBON_TITLED) : 0.75) : 1) : autoHeightFactor(g, true, g.type.equals("surface"), ns);
            double r = hPercent != null ? hPercent / 100 : plot.getHeight() / Math.max(1, plot.getWidth()) * factor;
            double w = horiz ? Math.min(plot.getWidth() / (1 + r * m / nCats * sy), plot.getHeight() / (r * (1 + m / nCats * sx)))
                    : Math.min(plot.getWidth() / (1 + m / nCats * sy), plot.getHeight() / (r + m / nCats * sx));
            double h = w * r, dd = (horiz ? h : w) / nCats * m;
            d3x = dd * sy;
            d3y = -dd * sx;
            double ox = plot.getX() + (plot.getWidth() - w - d3x) / 2, oy = plot.getY() + (plot.getHeight() - h + d3y) / 2 - d3y;
            plot = new Rectangle2D.Double(ox, oy, w, h);
            // 壁と床の塗り (横の壁は 0.63 倍、床は 0.75 倍に暗くする。横棒は床が左の面、横の壁が下の面)
            double fx = plot.getX(), fy = plot.getY(), fr = plot.getMaxX(), fb = plot.getMaxY();
            Shape back = new Rectangle2D.Double(fx + d3x, fy + d3y, w, h);
            Shape left = poly(fx, fy, fx + d3x, fy + d3y, fx + d3x, fb + d3y, fx, fb);
            Shape bottom = poly(fx, fb, fx + d3x, fb + d3y, fr + d3x, fb + d3y, fr, fb);
            obliqueWall(cv, "backWall", back, 1);
            obliqueWall(cv, "sideWall", horiz ? bottom : left, OBLIQUE_SIDE_SHADE);
            obliqueWall(cv, "floor", horiz ? left : bottom, OBLIQUE_FLOOR_SHADE);
            // 系列が奥行きに並ぶ 3D は、目盛りを粗くする (Excel の印刷: 0〜15 を 5 刻み。exp50・52 の 2 系列は 0〜12 の 2 刻みでなく 0〜15)
            autoScale(valAx, dmin, dmax, horiz ? plot.getWidth() : plot.getHeight(), ribbon ? vgap * RIBBON_VGAP : vgap);
        }
        dumpPlot = plot;
        dumpVal = valAx;
        // 補助目盛線 (主目盛線の下)
        if (valAx.gridMinor && !valAx.logScale) {
            DrawingPainter.LineStyle mg = valAx.gridMinorSpPr != null && Xml.child(valAx.gridMinorSpPr, "ln") != null
                    ? dp.line(Xml.child(valAx.gridMinorSpPr, "ln"), null, null) : defaultLine(new Color(0xF2F2F2), 0.75);
            double mu = valAx.minorUnit != null && valAx.minorUnit > 0 ? valAx.minorUnit : valAx.unit / 5;
            if (mg != null && mu > 0 && (valAx.hi - valAx.lo) / mu < 1000) {
                for (double v = valAx.lo; v <= valAx.hi + mu * 1e-6; v += mu) {
                    double p = valPos(valAx, v, plot, horiz);
                    dp.strokeShape(cv, horiz ? new Line2D.Double(p, plot.getY(), p, plot.getMaxY()) : new Line2D.Double(plot.getX(), p, plot.getMaxX(), p), mg);
                }
            }
        }
        // 目盛線
        DrawingPainter.LineStyle grid = valAx.gridMajor ? dp.line(valAx.gridSpPr == null ? null : Xml.child(valAx.gridSpPr, "ln"), null, null) : null;
        if (valAx.gridMajor && grid == null && (valAx.gridSpPr == null || Xml.child(valAx.gridSpPr, "ln") == null)) grid = gridDefault();
        for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v = step(valAx, v)) {
            double p = valPos(valAx, v, plot, horiz);
            if (grid != null && d3x > 0 && horiz) {
                // 横棒: 床の斜めの線と、奥の壁の縦の目盛線
                dp.strokeShape(cv, new Line2D.Double(p, plot.getMaxY(), p + d3x, plot.getMaxY() + d3y), grid);
                dp.strokeShape(cv, new Line2D.Double(p + d3x, plot.getY() + d3y, p + d3x, plot.getMaxY() + d3y), grid);
            } else if (grid != null && d3x > 0) {
                // 奥の壁の目盛線と、左の壁の斜めの線
                dp.strokeShape(cv, new Line2D.Double(plot.getX(), p, plot.getX() + d3x, p + d3y), grid);
                dp.strokeShape(cv, new Line2D.Double(plot.getX() + d3x, p + d3y, plot.getMaxX() + d3x, p + d3y), grid);
            } else if (grid != null) {
                Line2D l = horiz ? new Line2D.Double(p, plot.getY(), p, plot.getMaxY()) : new Line2D.Double(plot.getX(), p, plot.getMaxX(), p);
                dp.strokeShape(cv, l, grid);
            }
            if (!valAx.deleted && !"none".equals(valAx.tickLblPos)) {
                String lab = axisLabel(valAx, v, srcFmt);
                if (horiz) text(cv, lab, p, plot.getMaxY() + 3, vt, 0, -1);
                else if (valAxisX >= plot.getMaxX() - 0.01) text(cv, lab, valAxisX + valLabelGap(vt), p, vt, -1, 0);
                else text(cv, lab, valAxisX - valLabelGap(vt) - (scatter ? tickOut(valAx) : 0) - (legacy ? LEGACY_VAL_LABEL_OUT : 0), p, vt, 1, 0);
            }
        }
        // 項目軸の目盛線: 項目の境 (midCat なら項目の位置) ごとの線 (Excel の印刷 exp79)
        if (!scatter && catAx.gridMajor && d3x == 0) {
            DrawingPainter.LineStyle cg = dp.line(catAx.gridSpPr == null ? null : Xml.child(catAx.gridSpPr, "ln"), null, null);
            if (cg == null && (catAx.gridSpPr == null || Xml.child(catAx.gridSpPr, "ln") == null)) cg = gridDefault();
            if (cg != null) {
                double len = horiz ? plot.getHeight() : plot.getWidth();
                int nt = midCat ? nCats : nCats + 1;
                double tstep = len / Math.max(1, midCat ? nCats - 1 : nCats);
                for (int k = 0; k < nt; k++) {
                    double p = (horiz ? plot.getMaxY() : plot.getX()) + (horiz ? -1 : 1) * k * tstep;
                    dp.strokeShape(cv, horiz ? new Line2D.Double(plot.getX(), p, plot.getMaxX(), p) : new Line2D.Double(p, plot.getY(), p, plot.getMaxY()), cg);
                }
            }
        }
        if (scatter) {
            DrawingPainter.LineStyle xg = xValAx.gridMajor ? gridDefault() : null;
            for (double v = xValAx.lo; v <= xValAx.hi + xValAx.unit * 1e-6; v = step(xValAx, v)) {
                double p = plot.getX() + (v - xValAx.lo) / (xValAx.hi - xValAx.lo) * plot.getWidth();
                if (xg != null) dp.strokeShape(cv, new Line2D.Double(p, plot.getY(), p, plot.getMaxY()), xg);
                double ly = (scatterZeroInside(scatter, xValAx, valAx) ? valPos(valAx, 0, plot, false) + 3 : plot.getMaxY() + 3) + tickOut(xValAx);
                if (!xValAx.deleted) text(cv, axisLabel(xValAx, v, "General"), p, ly, ct, 0, -1);
            }
        }

        // 第 2 軸の目盛線とラベル (横棒は上、それ以外は右)
        if (valAx2 != null) {
            DrawingPainter.LineStyle grid2 = valAx2.gridMajor ? dp.line(valAx2.gridSpPr == null ? null : Xml.child(valAx2.gridSpPr, "ln"), null, null) : null;
            if (valAx2.gridMajor && grid2 == null && (valAx2.gridSpPr == null || Xml.child(valAx2.gridSpPr, "ln") == null)) grid2 = gridDefault();
            for (double v = valAx2.lo; v <= valAx2.hi + valAx2.unit * 1e-6; v = step(valAx2, v)) {
                double p = valPos(valAx2, v, plot, horiz);
                if (grid2 != null) {
                    Line2D l = horiz ? new Line2D.Double(p, plot.getY(), p, plot.getMaxY()) : new Line2D.Double(plot.getX(), p, plot.getMaxX(), p);
                    dp.strokeShape(cv, l, grid2);
                }
                if (val2Visible) {
                    String lab = axisLabel(valAx2, v, srcFmt2);
                    if (horiz) text(cv, lab, p, plot.getY() - 3 - (modern() ? 2.4 : 0), vt2, 0, 1);
                    else text(cv, lab, plot.getMaxX() + valLabelGap(vt2) + (modern() ? SEC_LABEL_EXTRA : LEGACY_SEC_LABEL_EXTRA), p, vt2, -1, 0);
                }
            }
        }

        // 系列
        for (Group gg : groups) {
            Axis va = secondary.contains(gg) ? valAx2 : valAx;
            switch (gg.type) {
                case "bar" -> paintBars(cv, gg, va, plot, nCats);
                case "line" -> paintLines(cv, gg, va, plot, nCats, false);
                case "area" -> paintLines(cv, gg, va, plot, nCats, true);
                case "scatter" -> paintScatter(cv, gg, va, xValAx, plot);
                case "surface" -> paintSurface(cv, gg, va, plot, nCats);
                default -> { }
            }
            if (gg.type.equals("line")) paintLineExtras(cv, gg, va, plot, nCats);
        }
        // 系列を奥に並べる斜投影の 3D: 系列名を床の右の縁に沿って奥へ並べる。重なるなら一つおき (Excel の印刷 exp72 の円柱: A と D)
        if (ribbon && !horiz && d3x > 0) {
            Axis sa = null;
            for (Axis a : axes) if (a.kind.equals("serAx") && (g.axIds.isEmpty() || g.axIds.contains(a.id))) sa = a;
            if (sa != null && !sa.deleted && !"none".equals(sa.tickLblPos)) {
                TStyle st = style(Xml.child(sa.el, "txPr"));
                int ns = g.series.size();
                double dy = Math.abs(d3y) / Math.max(1, ns), lh = lineH(st);
                int every = Math.max(1, (int) Math.ceil(lh / Math.max(1e-6, dy)));
                for (int k = 0; k < ns; k += every) {
                    double f = (k + 0.5) / ns;
                    text(cv, g.series.get(k).name, plot.getMaxX() + d3x * f + SER_LABEL_GAP, plot.getMaxY() + d3y * f, st, -1, 0);
                }
            }
        }

        // 値の軸の線と目盛り
        for (Axis va : valAx2 == null ? List.of(valAx) : List.of(valAx, valAx2)) {
            if (va.deleted || scatter) continue;
            DrawingPainter.LineStyle vl = axisLine(va);
            if (vl == null) continue;
            boolean second = va == valAx2;
            double at = horiz ? (second ? plot.getY() : plot.getMaxY()) : (second ? plot.getMaxX() : valAxisX);
            Line2D l = horiz ? new Line2D.Double(plot.getX(), at, plot.getMaxX(), at) : new Line2D.Double(at, plot.getY(), at, plot.getMaxY());
            dp.strokeShape(cv, l, vl);
            for (double v = va.lo; v <= va.hi + va.unit * 1e-6; v = step(va, v)) ticks(cv, va, vl, valPos(va, v, plot, horiz), at, horiz, second);
        }

        // 散布図の軸の線と目盛り: y の軸は valAxisX (x = 0 で交わるならその位置)、x の軸は y = 0 で交わるならその位置、なければ下端
        if (scatter) {
            DrawingPainter.LineStyle yl = valAx.deleted ? null : axisLine(valAx);
            if (yl != null) {
                dp.strokeShape(cv, new Line2D.Double(valAxisX, plot.getY(), valAxisX, plot.getMaxY()), yl);
                for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v = step(valAx, v)) ticks(cv, valAx, yl, valPos(valAx, v, plot, false), valAxisX, false, false);
            }
            DrawingPainter.LineStyle xl = xValAx.deleted ? null : axisLine(xValAx);
            if (xl != null) {
                double ay = scatterZeroInside(true, xValAx, valAx) ? valPos(valAx, 0, plot, false) : plot.getMaxY();
                dp.strokeShape(cv, new Line2D.Double(plot.getX(), ay, plot.getMaxX(), ay), xl);
                for (double v = xValAx.lo; v <= xValAx.hi + xValAx.unit * 1e-6; v = step(xValAx, v))
                    ticks(cv, xValAx, xl, plot.getX() + (v - xValAx.lo) / (xValAx.hi - xValAx.lo) * plot.getWidth(), ay, true, false);
            }
        }

        boolean table = dTable != null && !horiz && !scatter;
        if (table) paintDataTable(cv, groups, plot, nCats, ct, catAx);
        // 項目軸の線とラベル
        if (!scatter && catAx != null) {
            DrawingPainter.LineStyle axLine = catAx.el != null && Xml.child(catAx.el, "spPr") != null
                    ? dp.line(Xml.child(Xml.child(catAx.el, "spPr"), "ln"), null, null) : gridDefault();
            if (catAx.el != null && Xml.child(catAx.el, "spPr") != null && Xml.child(Xml.child(catAx.el, "spPr"), "ln") == null)
                axLine = gridDefault();
            // 項目軸の位置: 項目軸の crosses (autoZero = 0、max、min) または crossesAt (値の軸の上の値)
            double crossV = "max".equals(catAx.crosses) ? valAx.hi : "min".equals(catAx.crosses) ? valAx.lo
                    : catAx.crossesAt != null ? catAx.crossesAt : 0;
            double zero = valPos(valAx, Math.max(valAx.lo, Math.min(valAx.hi, crossV)), plot, horiz);
            if (!catAx.deleted && axLine != null) {
                Line2D l = horiz ? new Line2D.Double(zero, plot.getY(), zero, plot.getMaxY()) : new Line2D.Double(plot.getX(), zero, plot.getMaxX(), zero);
                dp.strokeShape(cv, l, axLine);
                double len = horiz ? plot.getHeight() : plot.getWidth();
                int nt = midCat ? nCats : nCats + 1;
                double tstep = len / Math.max(1, midCat ? nCats - 1 : nCats);
                for (int k = 0; k < nt; k++) ticks(cv, catAx, axLine, (horiz ? plot.getMaxY() : plot.getX()) + (horiz ? -1 : 1) * k * tstep, zero, !horiz, false);
            }
            if (!table && !catAx.deleted && !"none".equals(catAx.tickLblPos)) {
                Series s0 = g.series.isEmpty() ? null : g.series.get(0);
                List<String> cats = s0 == null ? List.of() : s0.cats;
                for (int i = 0; i < nCats; i++) {
                    String lab = i < cats.size() ? cats.get(i) : String.valueOf(i + 1);
                    if (horiz) {
                        double cy = catCenter(i, nCats, plot, true, catAx.reverse);
                        text(cv, lab, plot.getX() - valLabelGap(ct) - (legacy ? LEGACY_VAL_LABEL_OUT : 0), cy, ct, 1, 0);
                    } else {
                        double cx = catCenter(i, nCats, plot, false, catAx.reverse);
                        double slot = plot.getWidth() / Math.max(1, nCats);
                        // 「軸の横」のラベルは項目軸の線の下 (負の値があれば 0 の線の下)、「下端」はプロットの下
                        double ly = ("low".equals(catAx.tickLblPos) ? plot.getMaxY() + 3 : zero + 3) + catLabelDrop(ct, cats);
                        paintCatLabel(cv, lab, cx, ly, slot, ct);
                    }
                }
            }
        }
        Axis leftTitleAx = horiz ? catAx : valAx, bottomTitleAx = scatter ? xValAx : horiz ? valAx : catAx;
        if (bottomTitleAx != null && bottomTitleAx.titleEl != null) {
            TStyle tt = titleSize(axisTitleStyle(bottomTitleAx));
            String s = axisTitleText(bottomTitleAx);
            if (!s.isEmpty()) text(cv, s, plot.getCenterX(), box.getMaxY() - AXIS_TITLE_BOTTOM, tt, 0, 1);
        }
        if (leftTitleAx != null && leftTitleAx.titleEl != null) {
            TStyle tt = titleSize(axisTitleStyle(leftTitleAx));
            String s = axisTitleText(leftTitleAx);
            if (!s.isEmpty() && horizontalTitle(leftTitleAx)) {
                text(cv, s, box.getX() + H_TITLE_X, plot.getCenterY(), tt, -1, 0); // 横書きの縦の軸のタイトル (exp89)
            } else if (!s.isEmpty()) {
                cv.save();
                java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(box.getX() + lineH(tt) / 2, plot.getCenterY());
                at.rotate(-Math.PI / 2);
                cv.transform(at);
                text(cv, s, 0, 0, tt, 0, 0);
                cv.restore();
            }
        }
        if (!horiz && valAx.dispUnitLbl && valAx.dispUnitName != null && !valAx.deleted) {
            // 表示単位のラベル (日本語の Excel の表記)。目盛ラベルの左に縦書き
            String lab = switch (valAx.dispUnitName) {
                case "hundreds" -> "百"; case "thousands" -> "千"; case "tenThousands" -> "万"; case "hundredThousands" -> "十万";
                case "millions" -> "百万"; case "tenMillions" -> "千万"; case "hundredMillions" -> "億"; case "billions" -> "十億";
                case "trillions" -> "兆"; default -> "";
            };
            TStyle tt = dispUnitStyle(valAx);
            cv.save();
            java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(box.getX() + textLineH(tt, List.of(lab)) / 2 + DISP_UNIT_PAD,
                    plot.getY() + textWidth(lab, tt) / 2);
            at.rotate(-Math.PI / 2);
            cv.transform(at);
            text(cv, lab, 0, 0, tt, 0, 0);
            cv.restore();
        }
        if (valAx2 != null && valAx2.titleEl != null) {
            TStyle tt = axisTitleStyle(valAx2);
            tt.size = Math.max(tt.size, 10);
            String s = axisTitleText(valAx2);
            if (!s.isEmpty()) {
                if (horiz) {
                    text(cv, s, plot.getCenterX(), box.getY() + 4, tt, 0, -1);
                } else {
                    cv.save();
                    java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(box.getMaxX() - 4 - lineH(tt) / 2, plot.getCenterY());
                    at.rotate(-Math.PI / 2);
                    cv.transform(at);
                    text(cv, s, 0, 0, tt, 0, 0);
                    cv.restore();
                }
            }
        }
    }

    /** 横の軸のタイトルの、プロットの下の場所の追加 (タイトルの行の高さに足す) と、枠の下端から基線までの距離 (pt)。 */
    static final double AXIS_TITLE_GAP = 6, AXIS_TITLE_BOTTOM = 4;
    /** 横書きの縦の軸のタイトル: 左端からの位置と、タイトルの幅に足す場所 (pt。Excel の印刷 exp89)。 */
    static final double H_TITLE_X = 4.8, H_TITLE_GAP = 11;
    static final double H_PLOT_TOP = 0.6;

    /** 軸のタイトルの文字の大きさ (10pt 未満は 10pt)。 */
    private static TStyle titleSize(TStyle t) {
        t.size = Math.max(t.size, 10);
        return t;
    }

    /** 軸のタイトルの文字列。文字列の指定 (tx) がなければ Excel の既定の「軸ラベル」(言語が日本語以外なら Axis Title)。 */
    private String axisTitleText(Axis a) {
        if (a == null || a.titleEl == null) return "";
        if (Xml.child(a.titleEl, "tx") == null) return chartJa() ? "軸ラベル" : "Axis Title";
        return richText(a.titleEl);
    }

    /** グラフの言語が日本語か (c:lang がなければ日本語とみなす)。 */
    private boolean chartJa() {
        Element root = chartEl == null ? null : (Element) chartEl.getParentNode();
        String lang = root == null ? null : Xml.attr(Xml.child(root, "lang"), "val", null);
        return lang == null || lang.startsWith("ja");
    }

    /** 縦の軸のタイトルを横書きにするか (bodyPr の rot が 0)。 */
    private static boolean horizontalTitle(Axis a) {
        Element bp = Xml.path(a.titleEl, "txPr", "bodyPr");
        if (bp == null) bp = Xml.path(a.titleEl, "tx", "rich", "bodyPr");
        return bp != null && "0".equals(bp.getAttribute("rot"));
    }

    /** 軸ラベルの書式 (旧来のスタイルの既定は太字)。 */
    private TStyle axisTitleStyle(Axis a) {
        TStyle t = new TStyle();
        t.bold = legacy;
        t.apply(defaultTxPr);
        t.apply(Xml.child(a.titleEl, "txPr"));
        return t;
    }

    /** グループの値の範囲 (最小, 最大)。積み上げは合計、100% 積み上げは 0〜1。 */
    private static double[] dataRange(List<Group> groups, int nCats) {
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
        return new double[] {dmin, dmax};
    }

    /** 値の軸の線 (spPr の ln。なければ旧来のスタイルは灰色の線、2013 以降は描かない)。 */
    private DrawingPainter.LineStyle axisLine(Axis a) {
        Element sp = a.el == null ? null : Xml.child(a.el, "spPr");
        if (sp != null && Xml.child(sp, "ln") != null) return dp.line(Xml.child(sp, "ln"), null, null);
        return legacy ? gridDefault() : null;
    }

    /**
     * 主目盛を 1 本描く。
     *
     * @param p      目盛りの位置 (軸に沿った座標)
     * @param at     軸の線の位置
     * @param alongX 軸が横向き (目盛りは縦の線)
     * @param flip   外側が上・右 (第 2 軸、横棒の項目軸)
     */
    private void ticks(PdfCanvas cv, Axis a, DrawingPainter.LineStyle ls, double p, double at, boolean alongX, boolean flip) throws IOException {
        double len = 3;
        double out = switch (a.majorTick) { case "out" -> len; case "cross" -> len; default -> 0; };
        double in = switch (a.majorTick) { case "in" -> len; case "cross" -> len; default -> 0; };
        if (out == 0 && in == 0) return;
        double sgn = flip ? -1 : 1;
        if (alongX) {
            // 下向きが外側 (y が大きい方)。第 2 軸 (上) は逆
            dp.strokeShape(cv, new Line2D.Double(p, at - sgn * in, p, at + sgn * out), ls);
        } else {
            // 左が外側。第 2 軸 (右) は逆
            dp.strokeShape(cv, new Line2D.Double(at + sgn * in, p, at - sgn * out, p), ls);
        }
    }

    /** 目盛りの外側の長さ (out・cross で 3pt。散布図のラベルをその分離す)。 */
    private static double tickOut(Axis a) {
        return "out".equals(a.majorTick) || "cross".equals(a.majorTick) ? 3 : 0;
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

    /** 斜めの項目名の右上の角の、目盛りからの横のずれ (pt。exp125 の 6 件で最良) */
    static final double CR_DX = -4;
    /** 項目名を斜め (-45 度) に描く (catLabelHeight で決める)。 */
    private boolean catRot;

    /**
     * 項目名を斜めにするか (Excel の印刷 exp125): どれかが項目の幅に収まらず、しかも空白で区切った語のどれかが幅を超える
     * (和文の続いた文字列は 1 語。語の途中では折り返さない)。語が幅に収まれば空白で折り返す。
     */
    private boolean catRotate(List<String> cats, TStyle t, double slot) {
        boolean over = false;
        for (String c : cats) if (c != null && textWidth(c, t) > slot + 1) over = true;
        if (!over) return false;
        for (String c : cats) if (c != null) for (String w : c.split(" +")) if (!w.isEmpty() && textWidth(w, t) > slot + 1) return true;
        return false;
    }

    /** 項目ラベルの高さ (折り返し込み)。 */
    private double catLabelHeight(Group g, Axis catAx, TStyle t, double slot) {
        Series s0 = g.series.isEmpty() ? null : g.series.get(0);
        catRot = s0 != null && catRotate(s0.cats, t, slot);
        if (catRot) {
            double mw = 0;
            for (String c : s0.cats) if (c != null) mw = Math.max(mw, textWidth(c, t));
            return (mw + textLineH(t, s0.cats)) * Math.sqrt(0.5);
        }
        int lines = 1;
        if (s0 != null) for (String c : s0.cats) {
            double w = textWidth(c, t);
            if (w > slot) lines = Math.max(lines, (int) Math.ceil(w / Math.max(slot, 1)));
        }
        return Math.min(lines, 3) * textLineH(t, s0 == null ? List.of() : s0.cats);
    }

    /**
     * 文字に使う書体で決まる行の高さ (Excel の印刷 exp72): 和文を含めば和文の行 (lineH + 0.464em)、欧文だけなら欧文の書体の
     * アセント + ディセント。
     */
    private double textLineH(TStyle t, Iterable<String> texts) {

        for (String x : texts) if (x != null && x.codePoints().anyMatch(FontManager::isCjk)) return lineH(t) + TITLE_CJK_EXTRA * t.size;
        FontManager.Spec sp = t.spec();
        return (sp.face.ascent + sp.face.descent) * t.size;
    }

    /** 項目ラベルを描く (項目の幅で折り返す)。 */
    private void paintCatLabel(PdfCanvas cv, String lab, double cx, double y, double slot, TStyle t) throws IOException {
        double w = textWidth(lab, t);
        if (catRot) {
            // 斜めの項目名: 文字の右上の角を目盛りの位置に置き、左下へ向けて 45 度で描く (exp125)
            cv.save();
            java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(cx + CR_DX, y);
            at.rotate(-Math.PI / 4);
            cv.transform(at);
            text(cv, lab, 0, 0, t, 1, -1);
            cv.restore();
            return;
        }
        if (w <= slot + 1) {
            text(cv, lab, cx, y, t, 0, -1);
            return;
        }
        // 項目の幅に収まるよう折り返す
        TextLayout.Run run = new TextLayout.Run(lab, t.spec(), t.size, t.color);
        List<TextLayout.Line> lines = TextLayout.layout(List.of(run), (float) Math.max(slot, t.size), true, false, ctx.fonts);
        double yy = y;
        // 行の間隔は場所の計算 (catLabelHeight) と同じ textLineH (欧文 9pt で 11.0pt。Excel の印刷 exp81)
        double pitch = modern() ? textLineH(t, List.of(lab)) : -1;
        for (TextLayout.Line l : lines) {
            TextLayout.drawLine(cv, l, cx - l.width / 2, yy + l.ascent, 0, 0);
            yy += pitch > 0 ? pitch : l.ascent + l.descent;
        }
    }

    /** 次の目盛りの値 (対数目盛は底を掛ける)。 */
    private static double step(Axis a, double v) {
        return a.logScale ? v * a.logBase : v + a.unit;
    }

    /** 値の位置 (pt)。 */
    private static double valPos(Axis a, double v, Rectangle2D plot, boolean horiz) {
        double f = a.logScale ? (v <= a.lo ? 0 : Math.log(v / a.lo) / Math.log(a.hi / a.lo)) : (v - a.lo) / (a.hi - a.lo);
        if (a.reverse) f = 1 - f;
        return horiz ? plot.getX() + f * plot.getWidth() : plot.getMaxY() - f * plot.getHeight();
    }

    /** 項目の中心の位置 (pt)。 */
    /** データテーブル (なければ null)。 */
    private Element dTable;
    /** データテーブルの行の上下の余白の合計 (pt)。 */
    static final double DT_PAD = 2.75;
    /** データテーブルの左端のグラフの内側からの余白、凡例の印の幅 (四角・線) (pt)。Excel の印刷 exp72。 */
    static final double DT_BOTTOM = 1.5, DT_INSET = 9.5, DT_KEY_BOX = 7, DT_KEY_LINE = 21;

    /** 系列が線の凡例の印を使う (折れ線・散布図の線・レーダー)。 */
    private static boolean lineKey(List<Group> groups, Series s) {
        for (Group g : groups) if (g.series.contains(s))
            return g.type.equals("line") || g.type.equals("scatter") && !g.scatterStyle.equals("marker") || g.type.equals("radar");
        return false;
    }

    /** データテーブルの行の高さ: 文字の書体の行 + 余白。 */
    private double dtRowH(List<Group> groups, TStyle t) {
        List<String> texts = new ArrayList<>();
        for (Series s : allSeries(groups)) { texts.add(s.name); texts.addAll(s.cats); }
        return textLineH(t, texts) + DT_PAD;
    }

    /** すべてのグループの系列 (順に)。 */
    private static List<Series> allSeries(List<Group> groups) {
        List<Series> out = new ArrayList<>();
        for (Group gg : groups) out.addAll(gg.series);
        return out;
    }

    /** データテーブルの左の列の幅 (凡例の印、系列名、余白)。 */
    private double dtLeftWidth(List<Group> groups, TStyle t) {
        double mw = 0;
        for (Series s : allSeries(groups)) mw = Math.max(mw, textWidth(s.name, t));
        boolean keys = Xml.boolAttr(Xml.child(dTable, "showKeys"), "val", false), line = false;
        for (Series s : allSeries(groups)) line |= lineKey(groups, s);
        return mw + 7 + (keys ? (line ? DT_KEY_LINE : DT_KEY_BOX) : 0);
    }

    /** データテーブルを描く: 項目名の行と、系列ごとの値の行。 */
    private void paintDataTable(PdfCanvas cv, List<Group> groups, Rectangle2D plot, int n, TStyle t, Axis catAx) throws IOException {
        TStyle ts = style(Xml.child(dTable, "txPr"));
        if (Xml.child(dTable, "txPr") == null) ts = t;
        double rowH = dtRowH(groups, ts);
        List<Series> ss = allSeries(groups);
        boolean hz = Xml.boolAttr(Xml.child(dTable, "showHorzBorder"), "val", false), vt = Xml.boolAttr(Xml.child(dTable, "showVertBorder"), "val", false);
        boolean ol = Xml.boolAttr(Xml.child(dTable, "showOutline"), "val", false), keys = Xml.boolAttr(Xml.child(dTable, "showKeys"), "val", false);
        Element sp = Xml.child(dTable, "spPr");
        DrawingPainter.LineStyle ls = sp != null && Xml.child(sp, "ln") != null ? dp.line(Xml.child(sp, "ln"), null, null) : gridDefault();
        double x0 = plot.getX(), y0 = plot.getMaxY(), w = plot.getWidth(), slot = w / Math.max(1, n);
        double leftW = dtLeftWidth(groups, t), xl = x0 - leftW;
        double y1 = y0 + rowH * (1 + ss.size());
        List<String> cats = ss.isEmpty() ? List.of() : ss.get(0).cats;
        for (int i = 0; i < n; i++) {
            String lab = i < cats.size() ? cats.get(i) : String.valueOf(i + 1);
            text(cv, lab, x0 + (i + 0.5) * slot, y0 + rowH / 2, ts, 0, 0);
        }
        for (int k = 0; k < ss.size(); k++) {
            Series s = ss.get(k);
            double yy = y0 + rowH * (k + 1);
            double tx = xl + 4;
            if (keys) {
                if (lineKey(groups, s)) {
                    cv.line(tx, yy + rowH / 2, tx + DT_KEY_LINE - 4, yy + rowH / 2, new PdfCanvas.Stroke(2.25f), s.color);
                    tx += DT_KEY_LINE - 1;
                } else {
                    cv.fillRect(tx, yy + rowH / 2 - 3.5, 7, 7, s.color);
                    tx += 9;
                }
            }
            text(cv, s.name, tx, yy + rowH / 2, ts, -1, 0);
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                text(cv, ctx.formatter.formatNumber(v, s.fmt == null ? "General" : s.fmt), x0 + (i + 0.5) * slot, yy + rowH / 2, ts, 0, 0);
            }
        }
        if (ls == null) return;
        if (vt) for (int i = 0; i <= n; i++) dp.strokeShape(cv, new Line2D.Double(x0 + i * slot, y0, x0 + i * slot, y1), ls);
        if (hz) for (int k = 1; k <= ss.size(); k++) dp.strokeShape(cv, new Line2D.Double(xl, y0 + rowH * k, x0 + w, y0 + rowH * k), ls);
        if (ol) {
            dp.strokeShape(cv, new Line2D.Double(xl, y0 + rowH, xl, y1), ls);
            dp.strokeShape(cv, new Line2D.Double(xl, y1, x0 + w, y1), ls);
            dp.strokeShape(cv, new Line2D.Double(x0 + w, y0, x0 + w, y1), ls);
        }
    }

    /** 3D の視点 (度) と奥行き (%)。 */
    private double rotX = 15, rotY = 20, depthPct = 100;
    /** 直交軸 (斜投影) で描く。false は回転と遠近で投影する。 */
    private boolean rAngAx = true;
    /** c:chart (3D の壁・床の書式を読む)。 */
    private Element chartEl;
    /** 遠近の視野 (ST_Perspective。1/2 度単位) と、高さ (底の幅に対する %。null は自動)。 */
    private double perspective = 30;
    private Double hPercent;
    /** 3D の縦棒の奥行きの向きと長さ (pt。2D なら 0)。 */
    private double d3x, d3y;
    /** 描いているグループが 3D (2D の配置の規則 (exp81) を当てない)。 */
    private boolean flat3D;

    /** 2D の書式のあるグラフの配置の規則 (Excel の印刷 exp81) を当てる。旧来のスタイルと 3D は従来のまま。 */
    private boolean modern() {
        return !legacy && !flat3D;
    }

    /** グラフ全体の矩形 (手動で動かしたデータラベルの位置の基準)。 */
    private Rectangle2D chartRect;
    /** 円グラフの今の扇の中心と半径と角度 (ラジアン)。手動で動かしたラベルと引き出し線に使う。 */
    private double pieCx, pieCy, pieR = -1, pieMid;

    /** グラフの軸 (凡例で等高線の区間を決めるのに使う)。 */
    private List<Axis> chartAxes = List.of();

    /** 空のセルの表示 (gap, zero, span)。 */
    private String blanks = "gap";

    /** 項目を目盛りの上に置く (値の軸の crossBetween が midCat)。 */
    private boolean midCat;
    /** 縦の値の軸の横の位置。 */
    private double valAxisX;

    private double catCenter(int i, int n, Rectangle2D plot, boolean horiz, boolean reverse) {
        if (midCat && n > 1) {
            double len = horiz ? plot.getHeight() : plot.getWidth();
            int k = reverse ? n - 1 - i : i;
            return horiz ? plot.getMaxY() - k * len / (n - 1) : plot.getX() + k * len / (n - 1);
        }
        double slot = (horiz ? plot.getHeight() : plot.getWidth()) / Math.max(1, n);
        int k = reverse ? n - 1 - i : i;
        if (horiz) return plot.getMaxY() - (k + 0.5) * slot; // 最初の項目は下
        return plot.getX() + (k + 0.5) * slot;
    }

    /** 棒グラフを描く。 */
    private void paintBars(PdfCanvas cv, Group g, Axis valAx, Rectangle2D plot, int n) throws IOException {
        if (g.is3D && d3x > 0 && "standard".equals(g.grouping) && !g.horizontal) { paintBars3DStandard(cv, g, valAx, plot, n); return; }
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
        // 区分線のための、棒の矩形と値の端の位置
        Rectangle2D[][] rects = new Rectangle2D[g.series.size()][n];
        double[][] ends = new double[g.series.size()][n];
        for (int si = 0; si < g.series.size(); si++) {
            Series s = g.series.get(si);
            if (d3x > 0 && noFill(s)) continue; // 3D の塗りなしの系列は描かない (Excel の印刷 exp73)
            DrawingPainter.LineStyle ls = s.spPr == null ? null : dp.line(Xml.child(s.spPr, "ln"), null, null);
            double[] centers = new double[n];
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
                if (dp.bw) p = bwFill(col);
                boolean invOutline = false;
                if (v < 0 && Xml.boolAttr(Xml.child(s.ser, "invertIfNegative"), "val", false)) {
                    // 負の値の反転: 拡張 (c14:invertSolidFillFmt) の色、なければ白で塗る
                    Element inv = null;
                    for (Element ex : Xml.children(Xml.child(s.ser, "extLst"), "ext"))
                        for (Element f : Xml.children(ex, "invertSolidFillFmt")) inv = Xml.child(f, "spPr");
                    Color ic = inv != null && DrawingPainter.fillElement(inv) != null ? dp.paint(DrawingPainter.fillElement(inv), null, null).color : null;
                    p = solid(ic != null ? ic : Color.WHITE);
                    invOutline = ic == null; // 色の指定がなければ白で塗り、黒の 0.75pt の枠を描く (Excel の印刷 exp72)
                }
                if (d3x > 0) {
                    // 3D の箱: 系列の奥行き (gapDepth で決まる) の中ほどに置き、上面と右の面は暗くする
                    double f = 1 / (1 + g.gapDepth / 100);
                    double bx = d3x * f, by = d3y * f, ox = d3x * (1 - f) / 2, oy = d3y * (1 - f) / 2;
                    Rectangle2D front = new Rectangle2D.Double(bar.getX() + ox, bar.getY() + oy, bar.getWidth(), bar.getHeight());
                    Color bcol = p.color != null ? p.color : col;
                    double x0 = front.getX(), y0 = front.getY(), x1 = front.getMaxX(), y1 = front.getMaxY();
                    String shape = Xml.attr(Xml.child(s.ser, "shape"), "val", g.shape);
                    if (!shape.equals("box")) {
                        paintShapedBar(cv, front, bx, by, bcol, horiz, shape, v < 0);
                        paintDataLabel(cv, g, s, i, v, horiz ? new double[] {Math.max(from, to), front.getCenterY()} : new double[] {front.getCenterX(), Math.min(from, to)}, front);
                        rects[si][i] = front;
                        ends[si][i] = to;
                        centers[i] = front.getCenterX();
                        continue;
                    }
                    cv.fill(poly(x0, y0, x1, y0, x1 + bx, y0 + by, x0 + bx, y0 + by), shade(bcol, 0.8));
                    cv.fill(poly(x1, y0, x1 + bx, y0 + by, x1 + bx, y1 + by, x1, y1), shade(bcol, 0.6));
                    bar = front;
                }
                dp.fillShape(cv, bar, p, null);
                if (invOutline) cv.stroke(bar, new PdfCanvas.Stroke(0.75f), Color.BLACK);
                else if (ls != null) dp.strokeShape(cv, bar, ls);
                else if (dp.bw) cv.stroke(bar, new PdfCanvas.Stroke(0.75f), Color.BLACK); // 白黒印刷: 白の棒も見えるように黒の枠
                rects[si][i] = bar;
                ends[si][i] = to;
                paintDataLabel(cv, g, s, i, v, horiz ? new double[] {Math.max(from, to), bar.getCenterY()} : new double[] {bar.getCenterX(), Math.min(from, to)}, bar);
                centers[i] = bar.getCenterX();
            }
            if (!horiz) paintErrorBars(cv, s, i -> centers[i], valAx, plot);
        }
        if (!horiz) for (Series s : g.series) paintTrendlines(cv, s, catXs(n, plot), valAx, plot);
        if (stacked && g.serLines != null && d3x == 0) paintSerLines(cv, g, rects, ends, horiz);
    }

    /**
     * 円柱・円錐・ピラミッドの 3D の棒 (斜投影)。front は手前の面の矩形、(bx, by) は奥行きの向き。
     * 円は奥行きの平行四辺形に内接する楕円にする。円錐・ピラミッドは値の端を頂点にする (ToMax も同じ形で近似)。
     */
    private void paintShapedBar(PdfCanvas cv, Rectangle2D front, double bx, double by, Color c, boolean horiz, String shape, boolean neg)
            throws IOException {
        double x0 = front.getX(), y0 = front.getY(), x1 = front.getMaxX(), y1 = front.getMaxY();
        // 底 (値 0 の側) と端 (値の側) の面の中心と、面の 2 つの軸
        double[] base, end, u, w = {bx / 2, by / 2};
        if (horiz) {
            double cy = (y0 + y1) / 2 + by / 2;
            base = new double[] {(neg ? x1 : x0) + bx / 2, cy};
            end = new double[] {(neg ? x0 : x1) + bx / 2, cy};
            u = new double[] {0, (y1 - y0) / 2};
        } else {
            double cx = (x0 + x1) / 2 + bx / 2;
            base = new double[] {cx, (neg ? y0 : y1) + by / 2};
            end = new double[] {cx, (neg ? y1 : y0) + by / 2};
            u = new double[] {(x1 - x0) / 2, 0};
        }
        boolean pyramid = shape.startsWith("pyramid"), cone = shape.startsWith("cone");
        List<double[]> pts = new ArrayList<>();
        List<double[]> baseRing = ring(base, u, w, pyramid), endRing = ring(end, u, w, pyramid);
        pts.addAll(baseRing);
        if (pyramid || cone) pts.add(end); else pts.addAll(endRing);
        Path2D hull = hull(pts);
        cv.fill(hull, c);
        if (pyramid) {
            // 見える側の面 (縦棒は右、横棒は上) を暗くする
            double[] a = baseRing.get(horiz ? 2 : 1), b = baseRing.get(horiz ? 3 : 2);
            cv.fill(poly(a[0], a[1], b[0], b[1], end[0], end[1]), shade(c, 0.75));
        } else if (!cone) {
            Path2D cap = new Path2D.Double();
            for (int k = 0; k < endRing.size(); k++) { double[] q = endRing.get(k); if (k == 0) cap.moveTo(q[0], q[1]); else cap.lineTo(q[0], q[1]); }
            cap.closePath();
            cv.fill(cap, shade(c, 0.8));
        }
    }

    /** 中心 c と 2 つの軸 u, w の楕円 (pyramid なら平行四辺形の 4 つの角)。 */
    private static List<double[]> ring(double[] c, double[] u, double[] w, boolean square) {
        List<double[]> r = new ArrayList<>();
        if (square) {
            int[][] sg = {{-1, -1}, {1, -1}, {1, 1}, {-1, 1}};
            for (int[] g : sg) r.add(new double[] {c[0] + g[0] * u[0] + g[1] * w[0], c[1] + g[0] * u[1] + g[1] * w[1]});
            return r;
        }
        for (int k = 0; k < 48; k++) {
            double t = 2 * Math.PI * k / 48;
            r.add(new double[] {c[0] + u[0] * Math.cos(t) + w[0] * Math.sin(t), c[1] + u[1] * Math.cos(t) + w[1] * Math.sin(t)});
        }
        return r;
    }

    /** 点の凸包 (Andrew の方法)。 */
    private static Path2D hull(List<double[]> in) {
        List<double[]> p = new ArrayList<>(in);
        p.sort((a, b) -> a[0] != b[0] ? Double.compare(a[0], b[0]) : Double.compare(a[1], b[1]));
        int n = p.size(), k = 0;
        double[][] h = new double[2 * n][];
        for (int i = 0; i < n; i++) {
            while (k >= 2 && cross(h[k - 2], h[k - 1], p.get(i)) <= 0) k--;
            h[k++] = p.get(i);
        }
        for (int i = n - 2, t = k + 1; i >= 0; i--) {
            while (k >= t && cross(h[k - 2], h[k - 1], p.get(i)) <= 0) k--;
            h[k++] = p.get(i);
        }
        Path2D r = new Path2D.Double();
        for (int i = 0; i < k - 1; i++) { if (i == 0) r.moveTo(h[i][0], h[i][1]); else r.lineTo(h[i][0], h[i][1]); }
        r.closePath();
        return r;
    }

    private static double cross(double[] o, double[] a, double[] b) {
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
    }

    /** 積み上げ棒の区分線: 隣の項目の棒の、同じ系列の値の端どうしを結ぶ。 */
    private void paintSerLines(PdfCanvas cv, Group g, Rectangle2D[][] rects, double[][] ends, boolean horiz) throws IOException {
        Element sp = Xml.child(g.serLines, "spPr");
        DrawingPainter.LineStyle ls = sp != null && Xml.child(sp, "ln") != null ? dp.line(Xml.child(sp, "ln"), null, null)
                : defaultLine(legacy ? Color.BLACK : new Color(0x7F7F7F), 0.75);
        if (ls == null) return;
        for (int si = 0; si < rects.length; si++) {
            for (int i = 0; i + 1 < rects[si].length; i++) {
                Rectangle2D a = rects[si][i], b = rects[si][i + 1];
                if (a == null || b == null) continue;
                Line2D l;
                if (!horiz) l = new Line2D.Double(a.getMaxX(), ends[si][i], b.getX(), ends[si][i + 1]);
                else if (b.getY() < a.getY()) l = new Line2D.Double(ends[si][i], a.getY(), ends[si][i + 1], b.getMaxY());
                else l = new Line2D.Double(ends[si][i], a.getMaxY(), ends[si][i + 1], b.getY());
                dp.strokeShape(cv, l, ls);
            }
        }
    }

    /** 値を軸の範囲に収める。 */
    private static double clamp(double v, Axis a) { return Math.max(a.lo, Math.min(a.hi, v)); }

    /** 折れ線グラフ (area なら面グラフ) を描く。 */
    private void paintLines(PdfCanvas cv, Group g, Axis valAx, Rectangle2D plot, int n, boolean area) throws IOException {
        if (g.is3D && d3x > 0 && "standard".equals(g.grouping)) { paintRibbons(cv, g, valAx, plot, n, area); return; }
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
                if (v == null && blanks.equals("zero")) v = 0.0;
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
                if (!started) continue;   // 値のない系列 (rand30 rd30)
                for (int i = n - 1; i >= 0; i--) if (base[i] != null) p.lineTo(base[i][0], base[i][1]);
                p.closePath();
                DrawingPainter.Paint fill = !dp.bw && s.spPr != null && DrawingPainter.fillElement(s.spPr) != null
                        ? dp.paint(DrawingPainter.fillElement(s.spPr), null, null) : bwFill(s.color);
                dp.fillShape(cv, p, fill, null);
                // 白黒印刷の面は、黒・白・灰色の塗りに黒の枠 (Excel の印刷 exp124 bw_area)
                DrawingPainter.LineStyle ls = dp.bw ? defaultLine(Color.BLACK, BW_AREA_LINE) : s.spPr == null ? null : dp.line(Xml.child(s.spPr, "ln"), null, null);
                if (ls != null) dp.strokeShape(cv, p, ls);
            } else {
                polys.add(pts);
                DrawingPainter.LineStyle ls = seriesLine(s);
                Path2D p = new Path2D.Double();
                // 連続した点ごとに線にする (値のない点で切れる)。平滑線は曲線でつなぐ
                List<double[]> run = new ArrayList<>();
                for (int i = 0; i <= pts.length; i++) {
                    double[] pt = i < pts.length ? pts[i] : null;
                    if (pt != null) { run.add(pt); continue; }
                    if (i < pts.length && blanks.equals("span")) continue;
                    appendLine(p, run, s.smooth);
                    run.clear();
                }
                if (ls != null) dp.strokeShape(cv, p, ls);
                if (g.markers) for (int i = 0; i < n; i++) if (pts[i] != null) marker(cv, s, si, pts[i][0], pts[i][1]);
                paintTrendlines(cv, s, catXs(n, plot), valAx, plot);
                final int nn = n;
                paintErrorBars(cv, s, i -> catCenter(i, nn, plot, false, false), valAx, plot);
            }
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                if (pts[i] != null && s.vals.get(i) != null)
                    paintDataLabel(cv, g, s, i, s.vals.get(i), pts[i], null);
            }
        }
    }

    /**
     * 点の並びを折れ線 (smooth なら曲線) としてパスに加える。曲線は各点を通る 3 次ベジエで、制御点は前後の点を結ぶ
     * 向きの 1/6 (Catmull-Rom と同じ。Excel の平滑線も点の間で行き過ぎる)。
     */
    private static void appendLine(Path2D p, List<double[]> pts, boolean smooth) {
        if (pts.isEmpty()) return;
        p.moveTo(pts.get(0)[0], pts.get(0)[1]);
        for (int i = 1; i < pts.size(); i++) {
            double[] b = pts.get(i - 1), c = pts.get(i);
            if (!smooth || pts.size() < 3) { p.lineTo(c[0], c[1]); continue; }
            double[] a = pts.get(Math.max(0, i - 2)), d = pts.get(Math.min(pts.size() - 1, i + 1));
            p.curveTo(b[0] + (c[0] - a[0]) / 6, b[1] + (c[1] - a[1]) / 6, c[0] - (d[0] - b[0]) / 6, c[1] - (d[1] - b[1]) / 6, c[0], c[1]);
        }
    }

    /**
     * 折れ線のグループの補助の線: 降下線 (各項目の最大の点から項目軸まで)、高低線 (各項目の最大から最小まで)、
     * ローソク (最初と最後の系列の間の棒。上がりは白、下がりは黒)。系列より先に描く。
     */
    private void paintLineExtras(PdfCanvas cv, Group g, Axis va, Rectangle2D plot, int n) throws IOException {
        if (g.hiLowLines == null && g.upDownBars == null && g.dropLines == null) return;
        if (g.series.isEmpty()) return;
        double zero = valPos(va, Math.max(va.lo, Math.min(va.hi, 0)), plot, false);
        for (int i = 0; i < n; i++) {
            double lo = Double.MAX_VALUE, hi = -Double.MAX_VALUE;
            for (Series s : g.series) {
                Double v = i < s.vals.size() ? s.vals.get(i) : null;
                if (v == null) continue;
                lo = Math.min(lo, v);
                hi = Math.max(hi, v);
            }
            if (lo > hi) continue;
            double x = catCenter(i, n, plot, false, false);
            if (g.dropLines != null) {
                dp.strokeShape(cv, new Line2D.Double(x, valPos(va, clamp(hi, va), plot, false), x, zero), extraLine(g.dropLines));
            }
            if (g.hiLowLines != null && hi > lo) {
                dp.strokeShape(cv, new Line2D.Double(x, valPos(va, clamp(hi, va), plot, false), x, valPos(va, clamp(lo, va), plot, false)),
                        extraLine(g.hiLowLines));
            }
        }
        if (g.upDownBars != null && g.series.size() >= 2) {
            Series op = g.series.get(0), cl = g.series.get(g.series.size() - 1);
            double gap = Xml.doubleAttr(Xml.child(g.upDownBars, "gapWidth"), "val", 150) / 100;
            double slot = plot.getWidth() / Math.max(1, n), bw = slot / (1 + gap);
            for (int i = 0; i < n; i++) {
                Double o = i < op.vals.size() ? op.vals.get(i) : null, c = i < cl.vals.size() ? cl.vals.get(i) : null;
                if (o == null || c == null) continue;
                double x = catCenter(i, n, plot, false, false);
                double y0 = valPos(va, clamp(Math.max(o, c), va), plot, false), y1 = valPos(va, clamp(Math.min(o, c), va), plot, false);
                Rectangle2D r = new Rectangle2D.Double(x - bw / 2, y0, bw, Math.max(0.5, y1 - y0));
                Element bar = Xml.child(g.upDownBars, c >= o ? "upBars" : "downBars");
                Element sp = bar == null ? null : Xml.child(bar, "spPr");
                DrawingPainter.Paint fill = sp != null && DrawingPainter.fillElement(sp) != null ? dp.paint(DrawingPainter.fillElement(sp), null, null)
                        : solid(c >= o ? Color.WHITE : new Color(0x404040));
                dp.fillShape(cv, r, fill, null);
                DrawingPainter.LineStyle ls = sp != null && Xml.child(sp, "ln") != null ? dp.line(Xml.child(sp, "ln"), null, null)
                        : defaultLine(legacy ? Color.BLACK : new Color(0x404040), 0.75);
                if (ls != null) dp.strokeShape(cv, r, ls);
            }
        }
    }

    /** 降下線・高低線・誤差範囲の線 (spPr。なければ黒または濃い灰色の 0.75pt)。 */
    private DrawingPainter.LineStyle extraLine(Element e) {
        Element sp = e == null ? null : Xml.child(e, "spPr");
        if (sp != null && Xml.child(sp, "ln") != null) return dp.line(Xml.child(sp, "ln"), null, null);
        return defaultLine(legacy ? Color.BLACK : new Color(0x595959), 0.75);
    }

    /** 誤差範囲の各点の [下端, 上端] (値のない点は null)。 */
    private static List<double[]> errorSpans(Series s, Element eb) {
        String bt = Xml.attr(Xml.child(eb, "errBarType"), "val", "both");
        String vt = Xml.attr(Xml.child(eb, "errValType"), "val", "fixedVal");
        double val = Xml.doubleAttr(Xml.child(eb, "val"), "val", 1);
        double mean = 0, sd = 0;
        int cnt = 0;
        for (Double v : s.vals) if (v != null) { mean += v; cnt++; }
        if (cnt > 0) mean /= cnt;
        for (Double v : s.vals) if (v != null) sd += (v - mean) * (v - mean);
        sd = cnt > 1 ? Math.sqrt(sd / (cnt - 1)) : 0;
        List<double[]> out = new ArrayList<>();
        for (Double v : s.vals) {
            if (v == null) { out.add(null); continue; }
            double e = switch (vt) {
                case "percentage" -> Math.abs(v) * val / 100;
                case "stdDev" -> sd * val;
                case "stdErr" -> cnt > 0 ? sd / Math.sqrt(cnt) : 0;
                default -> val;
            };
            // 標準偏差は平均を中心にする (Excel と同じ)
            double c = vt.equals("stdDev") ? mean : v;
            out.add(new double[] {bt.equals("plus") ? v : c - e, bt.equals("minus") ? v : c + e});
        }
        return out;
    }

    /** 系列の誤差範囲 (縦方向) を描く。at は各点の横の位置。 */
    private void paintErrorBars(PdfCanvas cv, Series s, java.util.function.IntToDoubleFunction at, Axis va, Rectangle2D plot) throws IOException {
        if (s.ser == null) return;
        for (Element eb : Xml.children(s.ser, "errBars")) {
            if ("x".equals(Xml.attr(Xml.child(eb, "errDir"), "val", "y"))) continue;
            String bt = Xml.attr(Xml.child(eb, "errBarType"), "val", "both");
            boolean cap = !Xml.boolAttr(Xml.child(eb, "noEndCap"), "val", false);
            DrawingPainter.LineStyle ls = extraLine(eb);
            List<double[]> spans = errorSpans(s, eb);
            for (int i = 0; i < spans.size(); i++) {
                if (spans.get(i) == null) continue;
                double lo = spans.get(i)[0], hi = spans.get(i)[1];
                double x = at.applyAsDouble(i);
                double y0 = valPos(va, clamp(hi, va), plot, false), y1 = valPos(va, clamp(lo, va), plot, false);
                dp.strokeShape(cv, new Line2D.Double(x, y0, x, y1), ls);
                if (cap) {
                    if (!bt.equals("minus")) dp.strokeShape(cv, new Line2D.Double(x - 3, y0, x + 3, y0), ls);
                    if (!bt.equals("plus")) dp.strokeShape(cv, new Line2D.Double(x - 3, y1, x + 3, y1), ls);
                }
            }
        }
    }

    /** 折れ線・棒の x の値 (項目の番号 1, 2, …) を横の位置にする。 */
    private java.util.function.DoubleUnaryOperator catXs(int n, Rectangle2D plot) {
        double c0 = catCenter(0, n, plot, false, false), c1 = n > 1 ? catCenter(1, n, plot, false, false) : c0 + 1;
        return x -> c0 + (x - 1) * (c1 - c0);
    }

    /**
     * 系列の近似曲線 (trendline) を描く: 線形、指数、対数、累乗、多項式、移動平均。数式と R² の表示も描く。
     * x の値は散布図なら x、それ以外は項目の番号 (1 始まり)。
     */
    private void paintTrendlines(PdfCanvas cv, Series s, java.util.function.DoubleUnaryOperator xpx, Axis ya, Rectangle2D plot)
            throws IOException {
        if (s.ser == null) return;
        List<Element> tls = Xml.children(s.ser, "trendline");
        if (tls.isEmpty()) return;
        List<double[]> d = new ArrayList<>();
        for (int i = 0; i < s.vals.size(); i++) {
            Double v = s.vals.get(i);
            if (v == null) continue;
            double x = !s.xs.isEmpty() && i < s.xs.size() && s.xs.get(i) != null ? s.xs.get(i) : i + 1;
            d.add(new double[] {x, v});
        }
        if (d.size() < 2) return;
        for (Element t : tls) {
            try {
                paintTrendline(cv, s, t, d, xpx, ya, plot);
            } catch (RuntimeException e) {
                Warnings.add("近似曲線を描けませんでした", e);
            }
        }
    }

    private void paintTrendline(PdfCanvas cv, Series s, Element t, List<double[]> d, java.util.function.DoubleUnaryOperator xpx,
                                Axis ya, Rectangle2D plot) throws IOException {
        String type = Xml.attr(Xml.child(t, "trendlineType"), "val", "linear");
        double fwd = Xml.doubleAttr(Xml.child(t, "forward"), "val", 0), bwd = Xml.doubleAttr(Xml.child(t, "backward"), "val", 0);
        double xmin = Double.MAX_VALUE, xmax = -Double.MAX_VALUE;
        for (double[] p : d) { xmin = Math.min(xmin, p[0]); xmax = Math.max(xmax, p[0]); }
        xmin -= bwd;
        xmax += fwd;
        Path2D path = new Path2D.Double();
        java.util.function.DoubleUnaryOperator f = null;
        String eq = null;
        double r2 = Double.NaN;
        if (type.equals("movingAvg")) {
            int per = Math.max(2, Xml.intAttr(Xml.child(t, "period"), "val", 2));
            boolean started = false;
            for (int i = per - 1; i < d.size(); i++) {
                double sum = 0;
                for (int k = 0; k < per; k++) sum += d.get(i - k)[1];
                double px = xpx.applyAsDouble(d.get(i)[0]), py = valPos(ya, sum / per, plot, false);
                if (!started) { path.moveTo(px, py); started = true; } else path.lineTo(px, py);
            }
        } else {
            boolean logx = type.equals("log") || type.equals("power"), logy = type.equals("exp") || type.equals("power");
            int order = type.equals("poly") ? Math.max(2, Math.min(6, Xml.intAttr(Xml.child(t, "order"), "val", 2))) : 1;
            List<double[]> q = new ArrayList<>();
            for (double[] p : d) {
                if ((logx && p[0] <= 0) || (logy && p[1] <= 0)) continue;
                q.add(new double[] {logx ? Math.log(p[0]) : p[0], logy ? Math.log(p[1]) : p[1]});
            }
            if (q.size() < order + 1) return;
            double[] c = polyFit(q, order);
            // 決定係数 (指数・累乗は対数をとった値で計算する。Excel と同じ)
            double mean = 0;
            for (double[] p : q) mean += p[1];
            mean /= q.size();
            double ssr = 0, sst = 0;
            for (double[] p : q) {
                double e = p[1] - polyVal(c, p[0]);
                ssr += e * e;
                sst += (p[1] - mean) * (p[1] - mean);
            }
            r2 = sst > 0 ? 1 - ssr / sst : 1;
            final double[] cc = c;
            f = x -> {
                double xx = logx ? Math.log(x) : x;
                double y = polyVal(cc, xx);
                return logy ? Math.exp(y) : y;
            };
            eq = equation(type, c);
            double lo = logx ? Math.max(xmin, 1e-9) : xmin;
            int steps = 80;
            boolean started = false;
            for (int i = 0; i <= steps; i++) {
                double x = lo + (xmax - lo) * i / steps;
                double y = f.applyAsDouble(x);
                if (!Double.isFinite(y)) { started = false; continue; }
                double px = xpx.applyAsDouble(x), py = valPos(ya, y, plot, false);
                if (!started) { path.moveTo(px, py); started = true; } else path.lineTo(px, py);
            }
        }
        DrawingPainter.LineStyle ls;
        Element sp = Xml.child(t, "spPr");
        if (sp != null && Xml.child(sp, "ln") != null) ls = dp.line(Xml.child(sp, "ln"), null, null);
        else if (legacy) ls = defaultLine(Color.BLACK, 0.75);
        else {
            ls = defaultLine(s.color, 1.5);
            ls.dash = new float[] {1.5f, 1.5f};
        }
        cv.save();
        cv.clipRect(plot.getX(), plot.getY(), plot.getWidth(), plot.getHeight());
        if (ls != null) dp.strokeShape(cv, path, ls);
        cv.restore();
        // 数式と R² (近似曲線の右端の上)
        boolean showEq = Xml.boolAttr(Xml.child(t, "dispEq"), "val", false), showR = Xml.boolAttr(Xml.child(t, "dispRSqr"), "val", false);
        if ((showEq && eq != null) || (showR && !Double.isNaN(r2))) {
            TStyle ts = style(Xml.path(t, "trendlineLbl", "txPr"));
            List<String> lines = new ArrayList<>();
            if (showEq && eq != null) lines.add(eq);
            if (showR && !Double.isNaN(r2)) lines.add("R² = " + coef(r2));
            double ex = xpx.applyAsDouble(xmax), ey = f == null ? plot.getY() : valPos(ya, f.applyAsDouble(xmax), plot, false);
            if (modern()) {
                // 各行を中央に揃え、全体の右端を終点の 9.4pt 左、1 行目の基線を終点の 2.75pt 上に置く。行の間隔は textLineH
                // (Excel の印刷 exp72 の trend_linear: 9pt で 11.0pt)
                double pitch = textLineH(ts, lines), mw = 0;
                for (String l : lines) mw = Math.max(mw, textWidth(l, ts));
                double cx = Math.min(ex, plot.getMaxX()) - TREND_LBL_DX - mw / 2;
                double b0 = Math.max(plot.getY() + ts.size, ey - TREND_LBL_DY - (lines.size() - 2) * pitch / 2);
                FontManager.Spec tsp = ts.spec();
                for (int i = 0; i < lines.size(); i++) text(cv, lines.get(i), cx, b0 + i * pitch + tsp.descent * ts.size, ts, 0, 1);
            } else {
                double y0 = Math.max(plot.getY(), ey - lineH(ts) * (lines.size() + 0.5));
                for (int i = 0; i < lines.size(); i++) text(cv, lines.get(i), Math.min(ex, plot.getMaxX()) - 4, y0 + i * lineH(ts), ts, 1, -1);
            }
        }
    }

    /** 系列の近似曲線 (移動平均を除く) の値の最小・最大 (なければ null)。 */
    private double[] trendRange(Series s) {
        if (s.ser == null) return null;
        double lo = Double.MAX_VALUE, hi = -Double.MAX_VALUE;
        for (Element t : Xml.children(s.ser, "trendline")) {
            String type = Xml.attr(Xml.child(t, "trendlineType"), "val", "linear");
            if (type.equals("movingAvg")) continue;
            boolean logx = type.equals("log") || type.equals("power"), logy = type.equals("exp") || type.equals("power");
            int order = type.equals("poly") ? Math.max(2, Math.min(6, Xml.intAttr(Xml.child(t, "order"), "val", 2))) : 1;
            List<double[]> q = new ArrayList<>();
            double xmin = Double.MAX_VALUE, xmax = -Double.MAX_VALUE;
            for (int i = 0; i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                double x = !s.xs.isEmpty() && i < s.xs.size() && s.xs.get(i) != null ? s.xs.get(i) : i + 1;
                xmin = Math.min(xmin, x); xmax = Math.max(xmax, x);
                if ((logx && x <= 0) || (logy && v <= 0)) continue;
                q.add(new double[] {logx ? Math.log(x) : x, logy ? Math.log(v) : v});
            }
            if (q.size() < order + 1) continue;
            xmin -= Xml.doubleAttr(Xml.child(t, "backward"), "val", 0);
            xmax += Xml.doubleAttr(Xml.child(t, "forward"), "val", 0);
            double[] c = polyFit(q, order);
            for (int k = 0; k <= 40; k++) {
                double x = xmin + (xmax - xmin) * k / 40;
                if (logx && x <= 0) continue;
                double y = polyVal(c, logx ? Math.log(x) : x);
                if (logy) y = Math.exp(y);
                if (!Double.isFinite(y)) continue;
                lo = Math.min(lo, y); hi = Math.max(hi, y);
            }
        }
        return lo > hi ? null : new double[] {lo, hi};
    }

    /** 最小二乗の多項式 (係数は低い次数から)。 */
    static double[] polyFit(List<double[]> pts, int order) {
        int m = order + 1;
        double[][] a = new double[m][m + 1];
        for (double[] p : pts) {
            double[] pw = new double[2 * m];
            pw[0] = 1;
            for (int k = 1; k < 2 * m; k++) pw[k] = pw[k - 1] * p[0];
            for (int r = 0; r < m; r++) {
                for (int c = 0; c < m; c++) a[r][c] += pw[r + c];
                a[r][m] += pw[r] * p[1];
            }
        }
        // ガウスの消去法 (部分ピボット)
        for (int col = 0; col < m; col++) {
            int piv = col;
            for (int r = col + 1; r < m; r++) if (Math.abs(a[r][col]) > Math.abs(a[piv][col])) piv = r;
            double[] tmp = a[col]; a[col] = a[piv]; a[piv] = tmp;
            if (Math.abs(a[col][col]) < 1e-300) continue;
            for (int r = 0; r < m; r++) {
                if (r == col) continue;
                double k = a[r][col] / a[col][col];
                for (int c = col; c <= m; c++) a[r][c] -= k * a[col][c];
            }
        }
        double[] out = new double[m];
        for (int r = 0; r < m; r++) out[r] = Math.abs(a[r][r]) < 1e-300 ? 0 : a[r][m] / a[r][r];
        return out;
    }

    static double polyVal(double[] c, double x) {
        double y = 0;
        for (int k = c.length - 1; k >= 0; k--) y = y * x + c[k];
        return y;
    }

    /** 近似曲線の数式 (Excel の表記)。 */
    private String equation(String type, double[] c) {
        return switch (type) {
            case "exp" -> "y = " + coef(Math.exp(c[0])) + "e" + coef(c[1]) + "x";
            case "power" -> "y = " + coef(Math.exp(c[0])) + "x" + coef(c[1]);
            case "log" -> "y = " + coef(c[1]) + "ln(x)" + signed(c[0]);
            case "poly" -> {
                StringBuilder sb = new StringBuilder("y = ");
                for (int k = c.length - 1; k >= 1; k--) {
                    String v = k == c.length - 1 ? coef(c[k]) : signed(c[k]).trim();
                    if (k < c.length - 1) sb.append(' ');
                    sb.append(v).append('x');
                    if (k > 1) sb.append(k == 2 ? "²" : k == 3 ? "³" : String.valueOf(k));
                }
                yield sb.append(signed(c[0])).toString();
            }
            default -> "y = " + coef(c[1]) + "x" + signed(c[0]);
        };
    }

    /** 係数の表記 (小数 4 桁までの標準形式)。 */
    private String coef(double v) {
        double r = Math.round(v * 10000) / 10000.0;
        if (r == 0) r = 0;
        return ctx.formatter.formatNumber(r, "General");
    }

    /** 符号付きの項 (" + 1.5"、" - 6.5")。0 なら空。 */
    private String signed(double v) {
        double r = Math.round(v * 10000) / 10000.0;
        if (r == 0) return "";
        return (r < 0 ? " - " : " + ") + coef(Math.abs(r));
    }

    /**
     * 3D の折れ線・面: 系列ごとに奥行きの位置をずらし (最初の系列が手前)、奥の系列から描く。面は手前の面と上の帯、
     * 折れ線は帯 (リボン) で描く。
     */
    private void paintRibbons(PdfCanvas cv, Group g, Axis valAx, Rectangle2D plot, int n, boolean area) throws IOException {
        int ns = g.series.size();
        double sx = d3x / Math.max(1, ns), sy = d3y / Math.max(1, ns), f = 0.6;
        for (int k = ns - 1; k >= 0; k--) {
            Series s = g.series.get(k);
            double ox = sx * k, oy = sy * k;
            double[][] pts = new double[n][];
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                double x = area && n > 1 ? plot.getX() + i * plot.getWidth() / (n - 1) : catCenter(i, n, plot, false, false);
                pts[i] = new double[] {x + ox, valPos(valAx, clamp(v, valAx), plot, false) + oy};
            }
            Color c = s.color;
            if (s.spPr != null && DrawingPainter.fillElement(s.spPr) != null) {
                DrawingPainter.Paint pp = dp.paint(DrawingPainter.fillElement(s.spPr), null, null);
                if (pp != null && pp.color != null) c = pp.color;
            }
            // 上の帯 (各区間の四角形)
            for (int i = 0; i + 1 < n; i++) {
                if (pts[i] == null || pts[i + 1] == null) continue;
                double[] a = pts[i], b = pts[i + 1];
                cv.fill(poly(a[0], a[1], b[0], b[1], b[0] + sx * f, b[1] + sy * f, a[0] + sx * f, a[1] + sy * f), shade(c, area ? 0.8 : 1.0));
                if (!area) cv.stroke(poly(a[0], a[1], b[0], b[1], b[0] + sx * f, b[1] + sy * f, a[0] + sx * f, a[1] + sy * f), new PdfCanvas.Stroke(0.5f), shade(c, 0.7));
            }
            if (area) {
                double base = valPos(valAx, Math.max(valAx.lo, Math.min(valAx.hi, 0)), plot, false) + oy;
                Path2D p = new Path2D.Double();
                boolean started = false;
                for (double[] q : pts) {
                    if (q == null) continue;
                    if (!started) { p.moveTo(q[0], base); p.lineTo(q[0], q[1]); started = true; } else p.lineTo(q[0], q[1]);
                }
                for (int i = n - 1; i >= 0; i--) if (pts[i] != null) { p.lineTo(pts[i][0], base); break; }
                p.closePath();
                cv.fill(p, c);
            }
        }
    }

    /** 系列の線のスタイル。 */
    private DrawingPainter.LineStyle seriesLine(Series s) {
        DrawingPainter.LineStyle ls = null;
        if (s.spPr != null && Xml.child(s.spPr, "ln") != null) ls = dp.line(Xml.child(s.spPr, "ln"), null, null);
        else {
            ls = defaultLine(s.color, legacy ? 3 : 2.25);
            ls.cap = 1;
            ls.join = 1;
        }
        if (ls != null && ls.color == null && s.spPr != null && Xml.child(Xml.child(s.spPr, "ln"), "noFill") == null) ls.color = s.color;
        if (ls != null && dp.bw) ls.color = Color.BLACK; // 白黒印刷: 線は黒
        return ls;
    }

    /** マーカーを描く。 */
    private void marker(PdfCanvas cv, Series s, int si, double x, double y) throws IOException {
        String sym = "auto";
        double size = legacy ? 7 : 5;
        Color fill = s.color, line = s.color;
        if (s.marker != null) {
            sym = Xml.attr(Xml.child(s.marker, "symbol"), "val", "auto");
            size = Xml.doubleAttr(Xml.child(s.marker, "size"), "val", size);
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
            sym = auto[Math.max(0, s.idx) % auto.length]; // 系列の番号 (c:idx) で決まる
        }
        double h = legacy ? size * 9 / 14 : size * 0.75 / 2 + 1;
        Shape sh = switch (sym) {
            case "circle" -> new Ellipse2D.Double(x - h, y - h, 2 * h, 2 * h);
            case "dot" -> new Rectangle2D.Double(x - h * 0.8, y - h * 0.25, h * 1.6, h * 0.5); // 短い横棒
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

    /** 3D の縦棒 (系列ごとに奥行きの位置をずらす)。奥の系列から、箱 (前面・上面・右の面) で描く。 */
    private void paintBars3DStandard(PdfCanvas cv, Group g, Axis valAx, Rectangle2D plot, int n) throws IOException {
        int ns = g.series.size();
        double sx = d3x / Math.max(1, ns), sy = d3y / Math.max(1, ns);
        double slot = plot.getWidth() / Math.max(1, n), bw = slot / (1 + g.gapWidth / 100);
        double f = 1 / (1 + g.gapDepth / 100), bx = sx * f, by = sy * f;
        double zero = valPos(valAx, Math.max(valAx.lo, Math.min(valAx.hi, 0)), plot, false);
        for (int k = ns - 1; k >= 0; k--) {
            Series s = g.series.get(k);
            if (noFill(s)) continue;
            double ox = sx * k + sx * (1 - f) / 2, oy = sy * k + sy * (1 - f) / 2;
            for (int i = 0; i < n && i < s.vals.size(); i++) {
                Double v = s.vals.get(i);
                if (v == null) continue;
                double to = valPos(valAx, clamp(v, valAx), plot, false);
                double x0 = catCenter(i, n, plot, false, false) - bw / 2 + ox, y0 = Math.min(zero, to) + oy, x1 = x0 + bw, y1 = Math.max(zero, to) + oy;
                Color c = pointColor(g, s, i);
                String shape = Xml.attr(Xml.child(s.ser, "shape"), "val", g.shape);
                if (!shape.equals("box")) { paintShapedBar(cv, new Rectangle2D.Double(x0, y0, bw, y1 - y0), bx, by, c, false, shape, v < 0); continue; }
                cv.fill(poly(x0, y0, x1, y0, x1 + bx, y0 + by, x0 + bx, y0 + by), shade(c, 0.8));
                cv.fill(poly(x1, y0, x1 + bx, y0 + by, x1 + bx, y1 + by, x1, y1), shade(c, 0.6));
                cv.fill(new Rectangle2D.Double(x0, y0, bw, y1 - y0), c);
            }
        }
    }

    /**
     * 等高線グラフ: 値の目盛りの区間ごとに色を変える (区間 k は accent の k 番目)。3D は系列を奥行き方向に並べた面を
     * 奥の系列から四角形で描き、2D (上から見た図) は格子の各区画を中心の値の色で塗る。
     */
    private void paintSurface(PdfCanvas cv, Group g, Axis va, Rectangle2D plot, int n) throws IOException {
        int ns = g.series.size();
        if (ns == 0 || n < 2) return;
        java.util.function.DoubleFunction<Color> band = v -> autoColor((int) Math.max(0, Math.min(40, Math.floor((v - va.lo) / va.unit + 1e-9))));
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.5f);
        if (g.is3D && d3x > 0) {
            double sx = d3x / Math.max(1, ns - 1 > 0 ? ns - 1 : 1), sy = d3y / Math.max(1, ns - 1 > 0 ? ns - 1 : 1);
            for (int k = ns - 2; k >= 0; k--) {
                Series a = g.series.get(k), b = g.series.get(k + 1);
                for (int i = 0; i + 1 < n; i++) {
                    Double a0 = val(a, i), a1 = val(a, i + 1), b0 = val(b, i), b1 = val(b, i + 1);
                    if (a0 == null || a1 == null || b0 == null || b1 == null) continue;
                    double x0 = plot.getX() + i * plot.getWidth() / (n - 1), x1 = plot.getX() + (i + 1) * plot.getWidth() / (n - 1);
                    Path2D q = poly(x0 + sx * k, valPos(va, clamp(a0, va), plot, false) + sy * k,
                            x1 + sx * k, valPos(va, clamp(a1, va), plot, false) + sy * k,
                            x1 + sx * (k + 1), valPos(va, clamp(b1, va), plot, false) + sy * (k + 1),
                            x0 + sx * (k + 1), valPos(va, clamp(b0, va), plot, false) + sy * (k + 1));
                    Color c = band.apply((a0 + a1 + b0 + b1) / 4);
                    if (!g.wireframe) cv.fill(q, c);
                    cv.stroke(q, st, g.wireframe ? c : shade(c, 0.8));
                }
            }
        } else {
            double cw = plot.getWidth() / (n - 1), ch = plot.getHeight() / Math.max(1, ns - 1);
            for (int k = 0; k + 1 < Math.max(2, ns); k++) {
                Series a = g.series.get(Math.min(k, ns - 1)), b = g.series.get(Math.min(k + 1, ns - 1));
                for (int i = 0; i + 1 < n; i++) {
                    Double a0 = val(a, i), a1 = val(a, i + 1), b0 = val(b, i), b1 = val(b, i + 1);
                    if (a0 == null || a1 == null || b0 == null || b1 == null) continue;
                    Rectangle2D r = new Rectangle2D.Double(plot.getX() + i * cw, plot.getMaxY() - (k + 1) * ch, cw, ch);
                    Color c = band.apply((a0 + a1 + b0 + b1) / 4);
                    if (!g.wireframe) cv.fill(r, c); else cv.stroke(r, st, c);
                }
            }
        }
    }

    // ------------------------------------------------------------------ 回転と遠近の 3D

    /**
     * 3D の投影: 箱 (幅 W = 1、高さ H、奥行き D) の中心で、縦の軸の周りに rotY、横の軸の周りに rotX 回し、
     * 遠近 (perspective / 2 度の視野) をかけて、プロットの矩形に合わせる。座標は x 右、y 上、z 奥。
     */
    private static final class Proj {
        double w = 1, h, d, cy, sy, cx, sx, eye, s = 1, ox, oy;

        Proj(double rotX, double rotY, double persp, double h, double d) {
            this.h = h;
            this.d = d;
            cy = Math.cos(Math.toRadians(rotY)); sy = Math.sin(Math.toRadians(rotY));
            cx = Math.cos(Math.toRadians(rotX)); sx = Math.sin(Math.toRadians(rotX));
            double r = Math.sqrt(1 + h * h + d * d) / 2;
            eye = persp <= 0 ? Double.POSITIVE_INFINITY : r / Math.tan(Math.toRadians(Math.min(120, persp / 2)) / 2);
        }

        /** 視点の座標 (右、上、奥)。 */
        double[] view(double x, double y, double z) {
            double xx = x - w / 2, yy = y - h / 2, zz = z - d / 2;
            double x1 = xx * cy + zz * sy, z1 = -xx * sy + zz * cy;
            return new double[] {x1, yy * cx + z1 * sx, z1 * cx - yy * sx};
        }

        /** 紙の上の位置 (pt) と奥行き。 */
        double[] p(double x, double y, double z) {
            double[] v = view(x, y, z);
            double f = Double.isInfinite(eye) ? 1 : eye / (eye + v[2]);
            return new double[] {ox + v[0] * f * s, oy - v[1] * f * s, v[2]};
        }

        /** 収めるときに箱の外側に足す厚み (紙の上の pt。横の壁・床・奥の壁の向き)。 */
        double thx, thy, thz;

        /** 8 つの角の範囲 (縮尺 1、原点 0)。 */
        double[] bounds() {
            double s0 = s, x0 = ox, y0 = oy;
            double k = s0 > 0 ? 1 / s0 : 0, ex = thx * k, ey = thy * k, ez = thz * k;
            boolean leftWall = sy >= 0;
            s = 1; ox = 0; oy = 0;
            double[] b = {Double.MAX_VALUE, Double.MAX_VALUE, -Double.MAX_VALUE, -Double.MAX_VALUE};
            for (int i = 0; i < 8; i++) {
                double bx = (i & 1) == 0 ? (leftWall ? -ex : 0) : (leftWall ? w : w + ex);
                double[] q = p(bx, (i >> 1 & 1) == 0 ? -ey : h, (i >> 2 & 1) == 0 ? 0 : d + ez);
                b[0] = Math.min(b[0], q[0]); b[1] = Math.min(b[1], q[1]); b[2] = Math.max(b[2], q[0]); b[3] = Math.max(b[3], q[1]);
            }
            s = s0; ox = x0; oy = y0;
            return b;
        }

        /**
         * 視点の距離を紙の上の pt (eyePt) で決めて、矩形に収める。距離は縮尺に対して一定なので、縮尺と交互に求める
         * (Excel の実測 exp66: 視点の距離 = プロットの高さ / 2 ÷ tan(perspective / 4 度)、縮尺によらず一定)。
         */
        void fitEyePt(Rectangle2D r, double eyePt) {
            if (Double.isInfinite(eye)) { for (int i = 0; i < 30; i++) fit(r); return; }
            fit(r);
            for (int i = 0; i < 30; i++) {
                eye = eyePt / Math.max(1e-9, s);
                fit(r);
            }
        }

        /** 矩形に収める。 */
        void fit(Rectangle2D r) {
            double[] b = bounds();
            s = Math.min(r.getWidth() / (b[2] - b[0]), r.getHeight() / (b[3] - b[1]));
            ox = r.getCenterX() - (b[0] + b[2]) / 2 * s;
            oy = r.getCenterY() - (b[1] + b[3]) / 2 * s;
        }

        /** 面が見えるか (外向きの法線 n、面の中の点 c)。 */
        boolean faces(double[] n, double[] c) {
            double[] nv = dir(n), cv = view(c[0], c[1], c[2]);
            if (Double.isInfinite(eye)) return nv[2] < 0;
            return -(nv[0] * cv[0] + nv[1] * cv[1] + nv[2] * (cv[2] + eye)) > 0;
        }

        /** 向きの回転 (平行移動なし)。 */
        double[] dir(double[] n) {
            double x1 = n[0] * cy + n[2] * sy, z1 = -n[0] * sy + n[2] * cy;
            return new double[] {x1, n[1] * cx + z1 * sx, z1 * cx - n[1] * sx};
        }

        /** 面の明るさ (視点を向く面ほど明るい)。 */
        double light(double[] n) {
            double[] v = dir(n);
            double len = Math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]);
            return 0.62 + 0.38 * Math.min(1, Math.abs(v[2]) / Math.max(1e-9, len));
        }

        Path2D path(List<double[]> pts) {
            Path2D q = new Path2D.Double();
            for (int i = 0; i < pts.size(); i++) {
                double[] a = p(pts.get(i)[0], pts.get(i)[1], pts.get(i)[2]);
                if (i == 0) q.moveTo(a[0], a[1]); else q.lineTo(a[0], a[1]);
            }
            q.closePath();
            return q;
        }

        double depth(double x, double y, double z) { return view(x, y, z)[2]; }
    }

    /** 奥行き = 項目の幅 × DEPTH_K × depthPercent (exp66: 項目 4 つ、depth 100% で幅の 0.25、300% で 0.73)。 */
    static final double DEPTH_K = 1.0;
    /** 遠近の 3D のプロットの余白 (左、右、上、下。グラフの内側の枠 (外から 7pt) からの pt)。 */
    static final double[] PERSP_INSET = {4.8, 6.6, 6.3, 7.5};
    /** 遠近の 3D の値の軸の目盛りの最小の間隔 (文字の大きさに対する倍率)。 */
    static final double PERSP_VGAP = 1.45;
    /** 自動の高さの補正 (系列を奥に並べないもの。Excel の印刷 exp66〜69: 1.015 倍)。 */
    static final double PERSP_AUTO_H = 1.015;
    /** 系列を奥に並べるラベルのない 3D で、収めるときに横の壁と床の外側に足す厚み (紙の上の pt)。 */
    static final double PERSP_HIDDEN_WALL = 4;
    /** 奥の壁の外側に足す厚み (紙の上の pt) = この値 × (sin |rotY| − sin 20°)。既定の 20° では 0 (Excel の印刷 exp66 の rot60・mix、exp67 の rot50)。 */
    static final double PERSP_BACK_WALL = 23;
    /** 遠近が既定の 30 より強いときの視点の距離の倍率。 */
    static final double PERSP_WIDE_EYE = 0.92;
    /** 遠近の 3D の系列名を一つおきにする間隔の閾値 (行の高さの倍率)。 */
    static final double PERSP_SER_SKIP = 2.5;

    /** 軸の目盛ラベルを表示するか。 */
    private static boolean shown(Axis a) {
        return a != null && !a.deleted && !"none".equals(a.tickLblPos);
    }

    /** 系列の軸のラベルを表示するか (系列を奥に並べる 3D)。 */
    private static boolean serLabels(Group g, List<Axis> axes) {
        for (Axis a : axes) if (a.kind.equals("serAx") && (g.axIds.isEmpty() || g.axIds.contains(a.id)) && !a.deleted && !"none".equals(a.tickLblPos)) return true;
        return false;
    }

    /** 四角形 A-B-C-E を中心の点 (4 つの平均) で 4 つの三角形に分ける。 */
    private static double[][][] centerTris(double[] A, double[] B, double[] C, double[] E) {
        double[] M = new double[4];
        for (int j = 0; j < 4; j++) M[j] = (A[j] + B[j] + C[j] + E[j]) / 4;
        return new double[][][] {{A, B, M}, {B, C, M}, {C, E, M}, {E, A, M}};
    }

    /** 上から見た図の区間の数の上限 (autoScale の長さ ÷ 間隔)。 */
    static final double TOP_VIEW_BANDS = 4.5;
    /** 上から見たワイヤーフレームの線の太さ (Excel は画像で 1 画素ほど。exp72)。 */
    static final double TOP_WIRE_W = 0.25;

    /** 上から見た図 (rotX 90 の等高線) か。 */
    private boolean topViewChart() {
        return rotX >= 89.5;
    }

    /**
     * 自動の高さの倍率 (Excel の実測 exp68〜70。系列を奥に並べないもの・横棒は 1)。縦棒は系列 2・3 で 0.75、4 で 1、
     * 折れ線は 0.75、面・等高線は系列 2 で 0.5、3 以上で 0.75。
     */
    private double autoHeightFactor(Group g, boolean byDepth, boolean surface, int ns) {
        if (!byDepth || ns <= 1) return 1;
        if (surface || g.type.equals("area")) return ns == 2 ? 0.5 : 0.75;
        if (g.type.equals("bar")) return ns >= 4 ? 1 : 0.75;
        return 0.75;
    }

    /** 描く面 (奥から順に並べる)。 */
    /** 3D のワイヤーフレームの格子の辺 a-b (各点は x, y, z, 値) を、値の区間の境で切って区間の色の線にする。 */
    private void wireBand(List<Face3> faces, Proj P, double[] a, double[] b, Axis va) {
        List<Double> ts = new ArrayList<>(List.of(0.0, 1.0));
        double lo = Math.min(a[3], b[3]), hi = Math.max(a[3], b[3]);
        for (int band = 1; band < 40; band++) {
            double v = va.lo + band * va.unit;
            if (v > va.hi + 1e-9) break;
            if (v > lo && v < hi) ts.add((v - a[3]) / (b[3] - a[3]));
        }
        java.util.Collections.sort(ts);
        for (int j = 0; j + 1 < ts.size(); j++) {
            double t0 = ts.get(j), t1 = ts.get(j + 1), tm = (t0 + t1) / 2;
            double vm = a[3] + (b[3] - a[3]) * tm;
            int band = (int) Math.max(0, Math.min(39, Math.floor((vm - va.lo) / va.unit + 1e-9)));
            double[] p0 = lerp3(a, b, t0), p1 = lerp3(a, b, t1), pm = lerp3(a, b, tm);
            faces.add(new Face3(List.of(p0, p1), autoColor(band), P.depth(pm[0], pm[1], pm[2]), true));
        }
    }

    /** ワイヤーフレームの線分 a-b。 */
    private void wireLine(List<Face3> faces, Proj P, double[] a, double[] b, Color c) {
        double[] m = lerp3(a, b, 0.5);
        faces.add(new Face3(List.of(new double[] {a[0], a[1], a[2]}, new double[] {b[0], b[1], b[2]}), wireColor(c), P.depth(m[0], m[1], m[2]), true));
    }

    /** ワイヤーフレームの三角形の中の、区間の境の値の等値線 (下の区間の色)。 */
    private void wireIso(List<Face3> faces, Proj P, double[][] tri, Axis va, boolean pale) {
        for (int band = 1; band < 40; band++) {
            double v = va.lo + band * va.unit;
            if (v > va.hi + 1e-9) break;
            List<double[]> pts = new ArrayList<>();
            for (int e = 0; e < 3; e++) {
                double[] a = tri[e], b = tri[(e + 1) % 3];
                if (a[3] == b[3] || (a[3] - v) * (b[3] - v) > 0) continue;
                double[] q = lerp3(a, b, (v - a[3]) / (b[3] - a[3]));
                boolean dup = false;
                for (double[] o : pts) dup |= Math.abs(o[0] - q[0]) + Math.abs(o[1] - q[1]) + Math.abs(o[2] - q[2]) < 1e-9;
                if (!dup) pts.add(q);
            }
            if (pts.size() != 2) continue;
            double[] m = lerp3(pts.get(0), pts.get(1), 0.5);
            faces.add(new Face3(pts, pale ? wireColor(autoColor(band - 1)) : autoColor(band - 1), P.depth(m[0], m[1], m[2]), true));
        }
    }

    private static double[] lerp3(double[] a, double[] b, double t) {
        return new double[] {a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t};
    }

    private record Face3(List<double[]> pts, Color color, double depth, boolean outline) {}

    private static double[] v3(double x, double y, double z) { return new double[] {x, y, z}; }

    /** 3 点の面の法線 (a→b × a→c)。 */
    private static double[] normal(double[] a, double[] b, double[] c) {
        double ux = b[0] - a[0], uy = b[1] - a[1], uz = b[2] - a[2], vx = c[0] - a[0], vy = c[1] - a[1], vz = c[2] - a[2];
        return new double[] {uy * vz - uz * vy, uz * vx - ux * vz, ux * vy - uy * vx};
    }

    private static double[] centroid(List<double[]> pts) {
        double[] c = new double[3];
        for (double[] q : pts) for (int j = 0; j < 3; j++) c[j] += q[j] / pts.size();
        return c;
    }

    /**
     * 回転と遠近の 3D を描く。縦棒・横棒は箱、面は厚みのある板、折れ線は帯、等高線は値の区間ごとに色を変えた面。
     * 系列を奥行きに並べるもの (standard の縦棒、折れ線・面・等高線) は系列ごとに奥行きの列を持つ。
     */
    private void paintPersp(PdfCanvas cv, Group g, Axis catAx, Axis valAx, Axis serAx, Rectangle2D box, double dmin, double dmax,
                            String srcFmt, TStyle vt, TStyle ct) throws IOException {
        boolean horiz = g.type.equals("bar") && g.horizontal;
        boolean surface = g.type.equals("surface");
        boolean byDepth = surface || !g.type.equals("bar") || "standard".equals(g.grouping);
        boolean stacked = g.type.equals("bar") && (g.grouping.equals("stacked") || g.grouping.equals("percentStacked"));
        boolean percent = g.grouping.equals("percentStacked");
        int ns = Math.max(1, g.series.size());
        int n = 0;
        for (Series s : g.series) n = Math.max(n, Math.max(s.vals.size(), s.cats.size()));
        if (n == 0) return;
        boolean edges = surface || g.type.equals("area");
        // 奥行き: 系列を並べるものは (項目の幅 × 系列の数)、それ以外は項目の幅 (depthPercent を掛ける)
        double slot = 1.0 / n;
        // 奥行きは項目の幅に比例する (横棒は項目が縦に並ぶので、高さに比例する)
        // 系列を奥に並べるもの: 系列の数 × 項目の幅。集合縦棒: 棒の幅 × (1 + 間隔) (exp68: 系列 2・項目 4 で 0.182)
        double gapW = g.type.equals("bar") ? g.gapWidth / 100 : 1.5;
        // 横棒も同じ (項目の幅は高さの向きに測るので、あとで高さを掛ける。exp69)
        // ラベルのある 3D の等高線は奥行きを 0.92 倍 (Excel の印刷 exp72 の surf3d・surf_wire。ラベルのないもの (exp69・70) は 1)
        boolean anyLabel = shown(valAx) || shown(catAx) || shown(serAx);
        double depthK = byDepth ? slot * ns * DEPTH_K * depthPct / 100 * (surface && rotX < 89.5 && anyLabel ? 0.92 : 1)
                : (stacked ? slot / (1 + gapW) : slot / (ns + gapW)) * (1 + gapW) * depthPct / 100;
        java.util.function.DoubleUnaryOperator depthOf = hh -> {
            double d0 = Math.max(0.05, depthK * (horiz ? hh : 1));
            return d0;
        };
        // 上から見た図 (2D の等高線: rotX 90)。値の軸・壁は描かない
        boolean topView = rotX >= 89.5;
        TStyle st = serAx != null ? style(Xml.child(serAx.el, "txPr")) : ct;
        boolean valVis = !topView && !valAx.deleted && !"none".equals(valAx.tickLblPos);
        boolean catVis = catAx != null && !catAx.deleted && !"none".equals(catAx.tickLblPos);
        boolean serVis = byDepth && serAx != null && !serAx.deleted && !"none".equals(serAx.tickLblPos);
        // ラベルの場所
        double mw = 0;
        if (valVis) for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v = step(valAx, v)) mw = Math.max(mw, textWidth(axisLabel(valAx, v, srcFmt), vt));
        double cw = 0;
        Series s0 = g.series.isEmpty() ? null : g.series.get(0);
        if (catVis && s0 != null) for (String c : s0.cats) cw = Math.max(cw, textWidth(c, ct));
        double sw = 0;
        if (serVis) for (Series s : g.series) sw = Math.max(sw, textWidth(s.name, st));
        double left = horiz ? (catVis ? cw + 6 : 0) : (valVis ? mw + (surface ? 10 : 6) : 0); // 等高線は値のラベルの右に 10pt (exp72)
        double bottom = horiz ? (valVis ? lineH(vt) + 5 : 0) : (catVis ? lineH(ct) + 5 : 0);
        double right = serVis ? sw * 1.5 - 4 : 0; // 系列名の右の場所 (Excel の印刷 exp72 の等高線・exp50)
        // 余白 (グラフの内側の枠から。ラベルの場所と大きい方): exp66・exp67 の実測に合わせた値
        left = Math.max(left, PERSP_INSET[0]); right = Math.max(right, PERSP_INSET[1]); bottom = Math.max(bottom, PERSP_INSET[3]);
        // 和文の題名の行の増分は、3D の等高線では空けない (Excel の印刷 exp72。縦棒などの遠近の 3D (exp50・52) と上から見た図は空ける)
        double tx = topView || !surface ? 0 : titleExtra;
        // ラベルのある 3D の等高線は 1.5pt 下 (Excel の印刷 exp72 の surf3d・surf_wire)
        double sdy = surface && !topView && (valVis || catVis || serVis) ? 1.5 : 0;
        Rectangle2D plot = new Rectangle2D.Double(box.getX() + left, box.getY() + PERSP_INSET[2] - tx + sdy, Math.max(10, box.getWidth() - left - right),
                Math.max(10, box.getHeight() - bottom - PERSP_INSET[2] + tx));
        // 高さ: 指定がなければ、投影した箱の縦横の比がプロットの比に合うように決める (Excel の「自動調整」)
        Proj pj;
        if (hPercent != null) pj = new Proj(rotX, rotY, perspective, hPercent / 100, depthOf.applyAsDouble(hPercent / 100));
        else if (topView) {
            // 上から見た図: 横 : 縦 = 項目の数 : 系列の数 の長方形を、プロットに収まる最大で中央に置く
            // (Excel の印刷 exp72 の 4×4 は正方形、exp52 chartmisc2 の 5×3 は 5:3)
            double ratio = (double) Math.max(1, n) / Math.max(1, ns);
            double tw = Math.min(plot.getWidth(), plot.getHeight() * ratio), th = tw / ratio;
            plot = new Rectangle2D.Double(plot.getCenterX() - tw / 2, plot.getCenterY() - th / 2, tw, th);
            pj = new Proj(rotX, rotY, 0, 0.3, 1 / ratio);
        } else {
            // 自動の高さ: 箱の高さ ÷ 幅 = プロットの高さ ÷ 幅 × 倍率 (exp66: 360×270 の枠で 0.709 = 237.5 / 335。回転・遠近によらない。
            // 横棒も同じ (exp69))。倍率は系列を奥に並べるものの実測 (exp68〜70)
            double hh = plot.getHeight() / plot.getWidth() * autoHeightFactor(g, byDepth, surface, ns) * (byDepth && ns > 1 ? 1 : PERSP_AUTO_H);
            pj = new Proj(rotX, rotY, perspective, hh, depthOf.applyAsDouble(hh));
        }
        if (!topView) {
            // 系列を奥に並べ、軸のラベルがないときは、横の壁と床の外側に 4pt の厚みがあるものとして収める
            // (Excel の印刷 exp68〜70。棒の系列 4 以上は除く。ラベルがあるとき (exp50・52・72) は足さない)
            boolean dep = byDepth && ns > 1 && !valVis && !catVis && !serVis && !(g.type.equals("bar") && ns >= 4);
            boolean depL = byDepth && ns > 1 && !dep && !surface;
            // ラベルのある、系列を奥に並べる棒・折れ線・面は、横の壁だけ 3pt (Excel の印刷 exp52)
            pj.thx = dep ? PERSP_HIDDEN_WALL : depL ? 3 : 0; pj.thy = dep ? PERSP_HIDDEN_WALL : 0; pj.thz = PERSP_BACK_WALL * Math.max(0, Math.sin(Math.toRadians(Math.abs(rotY))) - Math.sin(Math.toRadians(20)));
        }
        if (!topView && perspective > 0) pj.fitEyePt(plot, plot.getHeight() / 2 / Math.tan(Math.toRadians(perspective / 4)) * (perspective > 30 ? PERSP_WIDE_EYE : 1)); // 既定 (30) より遠近が強いと視点は 0.92 倍 (exp66 の persp80・mix)
        else if (!topView) pj.fitEyePt(plot, Double.POSITIVE_INFINITY);
        else pj.fit(plot);
        final Proj P = pj;
        double W = 1, H = P.h, D = P.d;
        // 値の軸の目盛り: 投影した値の軸の長さで決める
        double[] a0 = horiz ? P.p(0, 0, 0) : P.p(0, 0, 0), a1 = horiz ? P.p(W, 0, 0) : P.p(0, H, 0);
        double axLen = Math.hypot(a1[0] - a0[0], a1[1] - a0[1]);
        // 目盛りの最小の間隔: 遠近の 3D の縦の値の軸は文字の大きさの PERSP_VGAP 倍 (exp52 chart3d2: 間 12.4pt の 2 刻みは使わない。exp50 chart3d の面: 間 17.3pt の 5 刻みは使う)
        // 上から見た図は、区間が 4 つ以下になる最小の刻み (Excel の印刷 exp72: 10〜60 は 20 刻み、exp52: 2〜11 は 5 刻み)
        if (topView) autoScale(valAx, dmin, dmax, TOP_VIEW_BANDS, 1);
        else autoScale(valAx, dmin, dmax, axLen, horiz ? 36 : vt.size * PERSP_VGAP);
        java.util.function.DoubleUnaryOperator vf = v -> {
            double f = valAx.logScale ? (v <= valAx.lo ? 0 : Math.log(v / valAx.lo) / Math.log(valAx.hi / valAx.lo)) : (v - valAx.lo) / (valAx.hi - valAx.lo);
            f = Math.max(0, Math.min(1, f));
            return valAx.reverse ? 1 - f : f;
        };
        // 壁の目盛線と床の線
        DrawingPainter.LineStyle grid = valAx.gridMajor ? dp.line(valAx.gridSpPr == null ? null : Xml.child(valAx.gridSpPr, "ln"), null, null) : null;
        if (valAx.gridMajor && grid == null && (valAx.gridSpPr == null || Xml.child(valAx.gridSpPr, "ln") == null)) grid = gridDefault();
        boolean leftWall = Math.sin(Math.toRadians(rotY)) >= 0;
        double wx = leftWall ? 0 : W;
        // 壁と床の塗り (backWall・sideWall・floor の spPr の塗り。厚みは描かない)
        if (!topView) {
            // 横棒は床が左 (値の軸の 0 の面)、横の壁が下 (exp69)
            List<double[]> bottomQ = List.of(v3(0, 0, 0), v3(W, 0, 0), v3(W, 0, D), v3(0, 0, D));
            List<double[]> sideQ = List.of(v3(wx, 0, 0), v3(wx, 0, D), v3(wx, H, D), v3(wx, H, 0));
            wallFill(cv, P, "floor", horiz ? sideQ : bottomQ);
            wallFill(cv, P, "backWall", List.of(v3(0, 0, D), v3(W, 0, D), v3(W, H, D), v3(0, H, D)));
            wallFill(cv, P, "sideWall", horiz ? bottomQ : sideQ);
        }
        for (double v = valAx.lo; v <= valAx.hi + valAx.unit * 1e-6; v = step(valAx, v)) {
            double f = vf.applyAsDouble(v);
            if (grid != null && !topView) {
                if (horiz) {
                    seg(cv, P, v3(f * W, 0, 0), v3(f * W, 0, D), grid);
                    seg(cv, P, v3(f * W, 0, D), v3(f * W, H, D), grid);
                } else {
                    seg(cv, P, v3(wx, f * H, 0), v3(wx, f * H, D), grid);
                    seg(cv, P, v3(0, f * H, D), v3(W, f * H, D), grid);
                }
            }
            if (valVis) {
                String lab = axisLabel(valAx, v, srcFmt);
                if (horiz) { double[] q = P.p(f * W, 0, 0); text(cv, lab, q[0], q[1] + 3, vt, 0, -1); }
                else { double[] q = P.p(wx, f * H, 0); text(cv, lab, Math.min(q[0], P.p(0, f * H, 0)[0]) - 4 - (surface ? 5 : 0), q[1] + (surface ? 2 : 0), vt, 1, 0); } // 等高線の値のラベルは 5pt 左・2pt 下 (exp72)
            }
        }
        // 床の縁 (項目の軸と系列の軸の線)
        DrawingPainter.LineStyle axl = gridDefault();
        if (!topView) {
            seg(cv, P, v3(0, 0, 0), v3(W, 0, 0), axl);
            seg(cv, P, v3(leftWall ? W : 0, 0, 0), v3(leftWall ? W : 0, 0, D), axl);
        }
        // 項目名・系列名
        if (catVis && s0 != null) {
            for (int i = 0; i < n && i < s0.cats.size(); i++) {
                double c = edges ? (n > 1 ? (double) i / (n - 1) : 0.5) : (i + 0.5) / n;
                if (horiz) { double[] q = P.p(0, c * H, 0); text(cv, s0.cats.get(i), q[0] - 4, q[1], ct, 1, 0); }
                else { double[] q = P.p(c * W, 0, 0); text(cv, s0.cats.get(i), q[0], q[1] + 3, ct, 0, -1); }
            }
        }
        if (serVis) {
            // 隣の系列名との縦の間隔が行の高さ × 2.5 より狭ければ一つおき (Excel の印刷 exp52 の 3D 折れ線・等高線: 系列 1 と 3)
            int every = 1;
            if (ns > 1) {
                // 隣の系列名の文字の枠が重なるなら、重ならなくなるまで間を空ける (Excel の印刷 exp52: 系列 1 と 3、exp88: 2 系列は両方)
                double z0 = surface ? 0 : D * 0.5 / ns, z1 = surface ? D / (ns - 1) : D * 1.5 / ns;
                double[] q0 = P.p(leftWall ? W : 0, 0, z0), q1 = P.p(leftWall ? W : 0, 0, z1);
                double dx = Math.abs(q1[0] - q0[0]), dy = Math.abs(q1[1] - q0[1]), lh = lineH(st);
                double sw0 = 0;
                for (Series ss : g.series) sw0 = Math.max(sw0, textWidth(ss.name, st));
                while (every < ns && dy * every < lh && dx * every < sw0) every++;
            }
            for (int k = 0; k < g.series.size(); k += every) {
                double z = surface && ns > 1 ? D * k / (ns - 1) : D * (k + 0.5) / ns;
                double[] q = P.p(leftWall ? W : 0, 0, z);
                text(cv, g.series.get(k).name, q[0] + (leftWall ? 4 : -4), q[1], st, leftWall ? -1 : 1, 0);
            }
        }
        // 系列の面を集めて、奥から描く
        List<Face3> faces = new ArrayList<>();
        double rowD = byDepth ? D / ns : D;
        double thick = rowD / (1 + g.gapDepth / 100);
        if (g.type.equals("bar")) {
            double gap = g.gapWidth / 100;
            double barW = stacked || byDepth ? slot / (1 + gap) : slot / (ns + gap);
            double[] pos = new double[n], neg = new double[n], tot = new double[n];
            if (percent) for (Series s : g.series) for (int i = 0; i < n; i++) { Double v = val(s, i); if (v != null) tot[i] += Math.abs(v); }
            for (int k = 0; k < g.series.size(); k++) {
                Series s = g.series.get(k);
                double z0 = (byDepth ? k * rowD : 0) + (rowD - thick) / 2;
                for (int i = 0; i < n; i++) {
                    Double v = val(s, i);
                    if (v == null) continue;
                    if (percent) v = tot[i] == 0 ? 0 : v / tot[i];
                    double from = 0, to = v;
                    if (stacked) { from = v >= 0 ? pos[i] : neg[i]; to = from + v; if (v >= 0) pos[i] = to; else neg[i] = to; }
                    else from = Math.max(valAx.lo, Math.min(valAx.hi, 0));
                    double f0 = vf.applyAsDouble(from), f1 = vf.applyAsDouble(to);
                    double c0 = (i + 0.5) / n - (stacked || byDepth ? barW / 2 : barW * ns / 2 - k * barW);
                    double[] lo3, hi3;
                    if (horiz) { lo3 = v3(Math.min(f0, f1) * W, c0 * H, z0); hi3 = v3(Math.max(f0, f1) * W, (c0 + barW) * H, z0 + thick); }
                    else { lo3 = v3(c0 * W, Math.min(f0, f1) * H, z0); hi3 = v3((c0 + barW) * W, Math.max(f0, f1) * H, z0 + thick); }
                    if (s.spPr != null && Xml.child(s.spPr, "noFill") != null) continue;   // 塗りなしの系列は描かない
                    addBox(faces, P, lo3, hi3, pointColor(g, s, i));
                }
            }
        } else if (surface) {
            int nz = Math.max(2, ns);
            java.util.function.IntToDoubleFunction zk = k -> D * k / (nz - 1);
            for (int k = 0; k + 1 < nz; k++) {
                Series a = g.series.get(Math.min(k, ns - 1)), b = g.series.get(Math.min(k + 1, ns - 1));
                for (int i = 0; i + 1 < n; i++) {
                    Double q00 = val(a, i), q10 = val(a, i + 1), q01 = val(b, i), q11 = val(b, i + 1);
                    if (q00 == null || q10 == null || q01 == null || q11 == null) continue;
                    double x0 = (double) i / (n - 1) * W, x1 = (double) (i + 1) / (n - 1) * W;
                    double[] A = {x0, vf.applyAsDouble(q00) * H, zk.applyAsDouble(k), q00}, B = {x1, vf.applyAsDouble(q10) * H, zk.applyAsDouble(k), q10},
                            C = {x1, vf.applyAsDouble(q11) * H, zk.applyAsDouble(k + 1), q11}, E = {x0, vf.applyAsDouble(q01) * H, zk.applyAsDouble(k + 1), q01};
                    if (g.wireframe && !topView) {
                        // 3D のワイヤーフレーム (exp72): 格子の線を区間の境で切って区間の色で描き、等値線を足す (色は薄めない)
                        wireBand(faces, P, A, B, valAx);
                        wireBand(faces, P, A, E, valAx);
                        if (i + 2 == n) wireBand(faces, P, B, C, valAx);
                        if (k + 2 == nz) wireBand(faces, P, E, C, valAx);
                        for (double[][] tri : new double[][][] {{A, B, E}, {B, C, E}}) wireIso(faces, P, tri, valAx, false);
                        continue;
                    }
                    if (g.wireframe && topView) {
                        // ワイヤーフレーム (Excel の印刷 exp72 の埋め込み画像で確認): 格子の線は区間の境で切って区間の色、
                        // 区間の境の値の等値線は下の区間の色 (どちらも薄めない)。対角線は描かない。左と奥の縁は灰色の線を重ねる
                        wireBand(faces, P, A, B, valAx);
                        wireBand(faces, P, A, E, valAx);
                        if (i + 2 == n) wireBand(faces, P, B, C, valAx);
                        if (k + 2 == nz) wireBand(faces, P, E, C, valAx);
                        for (double[][] tri : new double[][][] {{A, B, E}, {B, C, E}}) wireIso(faces, P, tri, valAx, false);
                        continue;
                    }
                    // 四角形は B-E の対角線で 2 つの三角形に分ける (Excel の印刷 exp72: 区間の境が直線につながる)
                    double[][][] tris = new double[][][] {{A, B, E}, {B, C, E}};
                    for (double[][] tri : tris) {
                        double[] nrm = normal(tri[0], tri[1], tri[2]);
                        // 光は上の手前から (上を向く面はほぼ元の色、奥へ傾く面は暗い)
                        double len = Math.sqrt(nrm[0] * nrm[0] + nrm[1] * nrm[1] + nrm[2] * nrm[2]);
                        double up = len == 0 ? 1 : Math.abs(nrm[1] * 0.8 - Math.signum(nrm[1]) * nrm[2] * 0.6) / len;
                        double li = 0.68 + 0.32 * Math.min(1, up);
                        for (int band = 0; band < 40; band++) {
                            double blo = valAx.lo + band * valAx.unit, bhi = blo + valAx.unit;
                            if (blo > valAx.hi + 1e-9) break;
                            List<double[]> part = clipBand(List.of(tri), band == 0 ? Double.NEGATIVE_INFINITY : blo, bhi);
                            if (part.size() < 3) continue;
                            Color c = autoColor(band);
                            double[] cc = centroid(part);
                            faces.add(new Face3(part, g.wireframe ? wireColor(c) : shade(c, li), P.depth(cc[0], cc[1], cc[2]), g.wireframe));
                        }
                    }
                }
            }
        } else {
            boolean area = g.type.equals("area");
            for (int k = 0; k < g.series.size(); k++) {
                Series s = g.series.get(k);
                if (noFill(s) && (area || lineNone(s))) continue; // 塗りなし (折れ線は線もなし) の系列は描かない (Excel の印刷 exp68)
                Color c = s.color;
                if (s.spPr != null && DrawingPainter.fillElement(s.spPr) != null) {
                    DrawingPainter.Paint pp = dp.paint(DrawingPainter.fillElement(s.spPr), null, null);
                    if (pp != null && pp.color != null) c = pp.color;
                }
                if (!area) {
                    DrawingPainter.LineStyle ls = seriesLine(s);
                    if (ls != null && ls.color != null) c = ls.color;
                }
                double z0 = k * rowD + (rowD - thick) / 2, z1 = z0 + thick;
                double[] xs = new double[n], ys = new double[n];
                boolean[] ok = new boolean[n];
                for (int i = 0; i < n; i++) {
                    Double v = val(s, i);
                    if (v == null) continue;
                    ok[i] = true;
                    xs[i] = area ? (n > 1 ? (double) i / (n - 1) : 0.5) * W : (i + 0.5) / n * W;
                    ys[i] = vf.applyAsDouble(v) * H;
                }
                double base = vf.applyAsDouble(Math.max(valAx.lo, Math.min(valAx.hi, 0))) * H;
                List<Face3> slab = new ArrayList<>();
                for (int i = 0; i + 1 < n; i++) {
                    if (!ok[i] || !ok[i + 1]) continue;
                    List<double[]> top = List.of(v3(xs[i], ys[i], z0), v3(xs[i + 1], ys[i + 1], z0), v3(xs[i + 1], ys[i + 1], z1), v3(xs[i], ys[i], z1));
                    double[] nrm = normal(top.get(0), top.get(3), top.get(1)); // 上向き
                    if (!area || P.faces(nrm, centroid(top))) slab.add(new Face3(top, shade(c, P.light(nrm) * (area ? 0.85 : 1)), 0, false));
                }
                if (area) {
                    int first = -1, last = -1;
                    for (int i = 0; i < n; i++) if (ok[i]) { if (first < 0) first = i; last = i; }
                    if (first >= 0) {
                        // 端の面
                        for (int e : new int[] {first, last}) {
                            List<double[]> cap = List.of(v3(xs[e], base, z0), v3(xs[e], ys[e], z0), v3(xs[e], ys[e], z1), v3(xs[e], base, z1));
                            double[] nrm = v3(e == first ? -1 : 1, 0, 0);
                            if (P.faces(nrm, centroid(cap))) slab.add(0, new Face3(cap, shade(c, P.light(nrm) * 0.8), 0, false));
                        }
                        // 手前の面
                        List<double[]> front = new ArrayList<>();
                        front.add(v3(xs[first], base, z0));
                        for (int i = first; i <= last; i++) if (ok[i]) front.add(v3(xs[i], ys[i], z0));
                        front.add(v3(xs[last], base, z0));
                        if (P.faces(v3(0, 0, -1), centroid(front))) slab.add(new Face3(front, c, 0, false));
                    }
                }
                double dep = P.depth(W / 2, H / 2, (z0 + z1) / 2);
                for (Face3 f : slab) faces.add(new Face3(f.pts(), f.color(), dep, f.outline()));
            }
        }
        // 奥のものから描く (同じ奥行きは加えた順)
        List<Face3> order = new ArrayList<>(faces);
        order.sort((x, y) -> Double.compare(y.depth(), x.depth()));
        for (Face3 f : order) {
            Path2D q = P.path(f.pts());
            if (f.outline()) { cv.stroke(q, new PdfCanvas.Stroke((float) (topView ? TOP_WIRE_W : 0.75)), f.color()); continue; }
            if (f.color() != null) {
                cv.fill(q, f.color());
                if (surface) cv.stroke(q, new PdfCanvas.Stroke(1.0f), f.color()); // 三角形の継ぎ目を埋める
            }
        }
        if (surface && g.wireframe && topView) {
            // 上から見たワイヤーフレームの左の縁と奥の縁に灰色の線 (Excel の印刷 exp72)
            Color gc = new Color(0xD9D9D9);
            seg(cv, P, v3(0, 0, 0), v3(0, 0, D), defaultLine(gc, 0.75));
            seg(cv, P, v3(0, 0, D), v3(W, 0, D), defaultLine(gc, 0.75));
        }
    }

    /** 3D の箱の見える面を加える (奥行きは箱の中心)。 */
    /** 3D の壁・床 (name) の塗りを、投影した面 pts に塗る。塗りがなければ何もしない。 */
    /** 等高線のワイヤーフレームの線の色: 区間の色を白と半々に混ぜた薄い色 (Excel の印刷 exp72)。 */
    static Color wireColor(Color c) {
        return new Color((c.getRed() + 255) / 2, (c.getGreen() + 255) / 2, (c.getBlue() + 255) / 2);
    }

    /** 直交軸の 3D の横の壁と床の暗さ (Excel の印刷 exp73: 純色が 0.63 倍・0.75 倍)。 */
    static final double OBLIQUE_SIDE_SHADE = 0.63, OBLIQUE_FLOOR_SHADE = 0.75;
    /** 直交軸の 3D の外側の余白の最小 (左・右・上・下、pt)。 */
    static final double[] OBLIQUE_INSET = {4.8, 8.5, 6.5, 6.2};
    /** 直交軸の 3D でラベルがあるときに足す左・下の余白と、横棒の右の余白 (pt)。 */
    static final double OBLIQUE_LABEL_LEFT = 4, OBLIQUE_LABEL_BOTTOM = 3.5, OBLIQUE_BAR_RIGHT = 13;

    /** 直交軸の 3D の壁・床を塗る (塗りの指定があるときだけ)。 */
    private void obliqueWall(PdfCanvas cv, String name, Shape sh, double k) throws IOException {
        Element sp = Xml.child(Xml.child(chartEl, name), "spPr");
        Element fe = sp == null ? null : DrawingPainter.fillElement(sp);
        if (fe == null || "noFill".equals(fe.getLocalName())) return;
        DrawingPainter.Paint pt = dp.paint(fe, null, null);
        if (pt == null || pt.color == null) return;
        cv.fill(sh, dp.bw ? Color.WHITE : k == 1 ? pt.color : shade(pt.color, k));
    }

    private void wallFill(PdfCanvas cv, Proj P, String name, List<double[]> pts) throws IOException {
        Element sp = Xml.child(Xml.child(chartEl, name), "spPr");
        Element fe = sp == null ? null : DrawingPainter.fillElement(sp);
        if (fe == null || "noFill".equals(fe.getLocalName())) return;
        DrawingPainter.Paint pt = dp.paint(fe, null, null);
        if (pt == null || pt.color == null) return;
        // 横の壁は 0.64 倍、床は 0.757 倍に暗くする (Excel の印刷 exp66: 緑 → #00A300、青 → #0000C1)
        double k = name.equals("sideWall") ? PERSP_SIDE_SHADE : name.equals("floor") ? PERSP_FLOOR_SHADE : 1;
        cv.fill(P.path(pts), dp.bw ? Color.WHITE : k == 1 ? pt.color : shade(pt.color, k));
    }

    static final double PERSP_SIDE_SHADE = 0.64, PERSP_FLOOR_SHADE = 0.757;

    private static void addBox(List<Face3> faces, Proj P, double[] lo, double[] hi, Color c) {
        double x0 = lo[0], y0 = lo[1], z0 = lo[2], x1 = hi[0], y1 = hi[1], z1 = hi[2];
        double dep = P.depth((x0 + x1) / 2, (y0 + y1) / 2, (z0 + z1) / 2);
        double[][] ns = {{0, 0, -1}, {0, 0, 1}, {0, 1, 0}, {0, -1, 0}, {1, 0, 0}, {-1, 0, 0}};
        List<List<double[]>> fs = List.of(
                List.of(v3(x0, y0, z0), v3(x1, y0, z0), v3(x1, y1, z0), v3(x0, y1, z0)),
                List.of(v3(x0, y0, z1), v3(x1, y0, z1), v3(x1, y1, z1), v3(x0, y1, z1)),
                List.of(v3(x0, y1, z0), v3(x1, y1, z0), v3(x1, y1, z1), v3(x0, y1, z1)),
                List.of(v3(x0, y0, z0), v3(x1, y0, z0), v3(x1, y0, z1), v3(x0, y0, z1)),
                List.of(v3(x1, y0, z0), v3(x1, y1, z0), v3(x1, y1, z1), v3(x1, y0, z1)),
                List.of(v3(x0, y0, z0), v3(x0, y1, z0), v3(x0, y1, z1), v3(x0, y0, z1)));
        for (int i = 0; i < 6; i++) {
            if (!P.faces(ns[i], centroid(fs.get(i)))) continue;
            faces.add(new Face3(fs.get(i), i == 0 ? c : shade(c, P.light(ns[i]) * 0.92), dep, false));
        }
    }

    /** 3D の線分を描く。 */
    private void seg(PdfCanvas cv, Proj P, double[] a, double[] b, DrawingPainter.LineStyle ls) throws IOException {
        double[] p = P.p(a[0], a[1], a[2]), q = P.p(b[0], b[1], b[2]);
        dp.strokeShape(cv, new Line2D.Double(p[0], p[1], q[0], q[1]), ls);
    }

    /**
     * 多角形 (点は {x, y, z, 値}) のうち、値が [lo, hi) の部分。値は辺の上で線形に変わるとして切る (等高線の塗り分け)。
     */
    static List<double[]> clipBand(List<double[]> poly, double lo, double hi) {
        List<double[]> r = clipHalf(poly, lo, true);
        return r.size() < 3 ? r : clipHalf(r, hi, false);
    }

    private static List<double[]> clipHalf(List<double[]> in, double t, boolean above) {
        List<double[]> out = new ArrayList<>();
        if (Double.isInfinite(t)) return new ArrayList<>(in);
        int m = in.size();
        for (int i = 0; i < m; i++) {
            double[] p = in.get(i), q = in.get((i + 1) % m);
            boolean pi = above ? p[3] >= t : p[3] < t, qi = above ? q[3] >= t : q[3] < t;
            if (pi) out.add(p);
            if (pi != qi) {
                double u = (t - p[3]) / (q[3] - p[3]);
                double[] r = new double[p.length];
                for (int j = 0; j < p.length; j++) r[j] = p[j] + u * (q[j] - p[j]);
                out.add(r);
            }
        }
        return out;
    }

    private static Double val(Series s, int i) { return i < s.vals.size() ? s.vals.get(i) : null; }

    /** 色を暗くする (f = 1 で元の色)。 */
    private static Color shade(Color c, double f) {
        return new Color((int) Math.min(255, Math.round(c.getRed() * f)), (int) Math.min(255, Math.round(c.getGreen() * f)),
                (int) Math.min(255, Math.round(c.getBlue() * f)));
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
                appendLine(p, pts, s.smooth || g.scatterStyle.startsWith("smooth"));
                dp.strokeShape(cv, p, seriesLine(s));
            }
            if (g.bubble) paintBubbles(cv, g, s, xAx, yAx, plot);
            else for (double[] pt : pts) marker(cv, s, si, pt[0], pt[1]);
            final Axis xa = xAx;
            paintTrendlines(cv, s, x -> plot.getX() + (x - xa.lo) / (xa.hi - xa.lo) * plot.getWidth(), yAx, plot);
        }
    }

    /** バブルの大きさの倍率の式の係数。 */
    static final double BUBBLE_A = 1.13, BUBBLE_K = 340;

    /**
     * バブルを描く。直径は大きさの平方根に比例し (面積が大きさに比例。幅を表す設定なら大きさに比例)。
     */
    private void paintBubbles(PdfCanvas cv, Group g, Series s, Axis xAx, Axis yAx, Rectangle2D plot) throws IOException {
        double max = 0;
        for (Group gg : List.of(g)) for (Series ss : gg.series) for (Double z : ss.sizes) if (z != null) max = Math.max(max, Math.abs(z));
        if (max <= 0) return;
        // 最大のバブルの直径 = プロットの短い辺 × 1.13 × s / (s + 340) (s は倍率 %。Excel の印刷 exp71: 50% で 0.143、100% で 0.255、200% で 0.41)
        double sc = Math.max(0, g.bubbleScale);
        double maxD = Math.min(plot.getWidth(), plot.getHeight()) * BUBBLE_A * sc / (sc + BUBBLE_K);
        DrawingPainter.Paint fill = s.spPr != null && DrawingPainter.fillElement(s.spPr) != null ? dp.paint(DrawingPainter.fillElement(s.spPr), null, null) : solid(s.color);
        DrawingPainter.LineStyle ls = s.spPr == null ? null : dp.line(Xml.child(s.spPr, "ln"), null, null);
        for (int i = 0; i < s.vals.size() && i < s.sizes.size(); i++) {
            Double v = s.vals.get(i), z = s.sizes.get(i);
            if (v == null || z == null || z <= 0) continue;
            double xv = i < s.xs.size() && s.xs.get(i) != null ? s.xs.get(i) : i + 1;
            double x = plot.getX() + (xv - xAx.lo) / (xAx.hi - xAx.lo) * plot.getWidth(), y = valPos(yAx, v, plot, false);
            double d = g.bubbleWidth ? maxD * z / max : maxD * Math.sqrt(z / max);
            Ellipse2D e = new Ellipse2D.Double(x - d / 2, y - d / 2, d, d);
            Element pp = pointSpPr(s, i);
            DrawingPainter.Paint pf = pp != null && DrawingPainter.fillElement(pp) != null ? dp.paint(DrawingPainter.fillElement(pp), null, null) : fill;
            DrawingPainter.LineStyle pl = pp != null && Xml.child(pp, "ln") != null ? dp.line(Xml.child(pp, "ln"), null, null) : ls;
            dp.fillShape(cv, e, pf == null ? solid(s.color) : pf, null);
            if (pl != null) dp.strokeShape(cv, e, pl);
        }
    }

    /**
     * 補助円グラフ・補助縦棒付き円グラフ: 最後の splitPos 個の要素を右の補助に回し、主の円ではその合計を 1 つの扇 (3 時の方向が中心) にする。
     * 補助の円は最初の要素を 9 時の方向を中心に、時計回りに置く。主と補助を区分線でつなぐ。
     */
    /** 補助円グラフの余白 (pt): 幅から引く量、高さから引く量。 */
    private static final double OF_PIE_HPAD = 5.5, OF_PIE_VPAD = 7.7;

    private void paintOfPie(PdfCanvas cv, Group g, Rectangle2D box) throws IOException {
        Series s = g.series.get(0);
        int n = s.vals.size(), k = Math.max(0, Math.min(n, g.splitPos));
        double mainTot = 0, secTot = 0;
        for (int i = 0; i < n; i++) {
            Double v = s.vals.get(i);
            if (v == null || v <= 0) continue;
            if (i >= n - k) secTot += v; else mainTot += v;
        }
        double tot = mainTot + secTot;
        if (tot <= 0) return;
        boolean barOf = "bar".equals(g.ofPie);
        // 補助の円は直径 d × f2、補助の縦棒は高さ d × f2・幅はその半分。主との間隔は円で gapWidth/200 × d、縦棒で gapWidth/400 × d
        double f2 = g.secondPieSize / 100, gap = g.gapWidth / (barOf ? 400.0 : 200.0), sw = barOf ? f2 / 2 : f2;
        // 主の直径 d: 全体 (d + 間隔 + 補助の幅) がプロットの幅 − 5.5pt に、d がプロットの高さ − 6pt に収まる。左端はプロットの左 + 3pt (Excel の実測)
        double d = Math.max(0, Math.min(box.getHeight() - OF_PIE_VPAD, (box.getWidth() - OF_PIE_HPAD) / (1 + gap + sw)));
        double total = d * (1 + gap + sw);
        double cx = box.getCenterX() - total / 2 + d / 2, cy = box.getCenterY();
        double d2 = d * f2, cx2 = cx + d / 2 + d * gap + d * sw / 2;
        DrawingPainter.LineStyle edge = s.spPr != null && Xml.child(s.spPr, "ln") != null ? dp.line(Xml.child(s.spPr, "ln"), null, null)
                : defaultLine(dp.bw ? Color.BLACK : Color.WHITE, 0.75);
        double otherExt = -secTot / tot * 360;
        double start = -otherExt / 2;
        // その他の扇 (補助に回した要素の合計)
        Color otherColor = autoColor(n);
        Arc2D other = new Arc2D.Double(cx - d / 2, cy - d / 2, d, d, start, otherExt, Arc2D.PIE);
        if (secTot > 0) { cv.fill(other, otherColor); if (edge != null) dp.strokeShape(cv, other, edge); }
        double a = start + otherExt;
        for (int i = 0; i < n - k; i++) {
            Double v = s.vals.get(i);
            if (v == null || v <= 0) continue;
            double ext = -v / tot * 360;
            Arc2D sl = new Arc2D.Double(cx - d / 2, cy - d / 2, d, d, a, ext, Arc2D.PIE);
            cv.fill(sl, pointColor(g, s, i));
            if (edge != null) dp.strokeShape(cv, sl, edge);
            a += ext;
        }
        if (secTot <= 0) return;
        double top = Math.toRadians(start), bot = Math.toRadians(start + otherExt);
        double[] p0 = {cx + Math.cos(top) * d / 2, cy - Math.sin(top) * d / 2}, p1 = {cx + Math.cos(bot) * d / 2, cy - Math.sin(bot) * d / 2};
        double[] q0, q1;
        if (barOf) {
            double bw = d2 / 2, bh = d2, bx = cx2 - bw / 2, by = cy - bh / 2;
            double y = by;
            for (int i = n - k; i < n; i++) {
                Double v = s.vals.get(i);
                if (v == null || v <= 0) continue;
                double h = bh * v / secTot;
                Rectangle2D r = new Rectangle2D.Double(bx, y, bw, h);
                cv.fill(r, pointColor(g, s, i));
                if (edge != null) dp.strokeShape(cv, r, edge);
                y += h;
            }
            q0 = new double[] {bx, by};
            q1 = new double[] {bx, by + bh};
        } else {
            // 最初の要素は、主の円の「その他」の扇の下の辺と同じ角度から時計回りに置く
            double b = start + otherExt;
            for (int i = n - k; i < n; i++) {
                Double v = s.vals.get(i);
                if (v == null || v <= 0) continue;
                double ext = -v / secTot * 360;
                Arc2D sl = new Arc2D.Double(cx2 - d2 / 2, cy - d2 / 2, d2, d2, b, ext, Arc2D.PIE);
                cv.fill(sl, pointColor(g, s, i));
                if (edge != null) dp.strokeShape(cv, sl, edge);
                b += ext;
            }
            q0 = new double[] {cx2, cy - d2 / 2};
            q1 = new double[] {cx2, cy + d2 / 2};
        }
        DrawingPainter.LineStyle sl = g.serLines != null && Xml.child(g.serLines, "spPr") != null && Xml.child(Xml.child(g.serLines, "spPr"), "ln") != null
                ? dp.line(Xml.child(Xml.child(g.serLines, "spPr"), "ln"), null, null) : defaultLine(legacy ? Color.BLACK : new Color(0xA6A6A6), 0.75);
        if (sl != null) {
            dp.strokeShape(cv, new Line2D.Double(p0[0], p0[1], q0[0], q0[1]), sl);
            dp.strokeShape(cv, new Line2D.Double(p1[0], p1[1], q1[0], q1[1]), sl);
        }
    }

    /**
     * 3D の円グラフ: 上面は傾けた楕円 (縦の半径 = 横の半径 × sin(rotX))、厚さは横の半径の 22%。手前の半分の側面を先に描く。
     */
    private void paintPie3D(PdfCanvas cv, Group g, Rectangle2D box) throws IOException {
        Series s = g.series.get(0);
        double total = 0;
        for (Double v : s.vals) if (v != null && v > 0) total += v;
        if (total <= 0) return;
        if (!rAngAx && perspective > 0) { paintPie3DPersp(cv, g, s, total, box); return; }
        double sinx = Math.max(0.1, Math.sin(Math.toRadians(rotX)));
        // 上面の楕円と厚さが、矩形の高さの 92%・幅の 90% に収まる大きさ
        double rx = Math.min(box.getWidth() * 0.45, box.getHeight() * 0.92 / (2 * sinx + 0.22 * depthPct / 100)), ry = rx * sinx;
        double th = rx * 0.22 * depthPct / 100;
        double cx = box.getCenterX(), cy = box.getCenterY() - th / 2;
        double start = 90 - g.firstSliceAng - rotY;
        List<double[]> slices = new ArrayList<>();
        for (int i = 0; i < s.vals.size(); i++) {
            Double v = s.vals.get(i);
            if (v == null || v <= 0) continue;
            double ext = -v / total * 360;
            slices.add(new double[] {i, start, ext});
            start += ext;
        }
        DrawingPainter.LineStyle edge = defaultLine(dp.bw ? Color.BLACK : Color.WHITE, 0.75);
        // 側面 (手前の半分: 角度 180〜360 度)
        for (double[] sl : slices) {
            double a0 = sl[1] + sl[2], a1 = sl[1]; // a0 < a1 (時計回りの扇)
            double from = a0, to = a1;
            Path2D side = new Path2D.Double();
            boolean any = false;
            int steps = 48;
            List<double[]> pts = new ArrayList<>();
            for (int k = 0; k <= steps; k++) {
                double t = from + (to - from) * k / steps;
                double tn = norm(t);
                if (tn >= 180 && tn <= 360) { pts.add(new double[] {cx + rx * Math.cos(Math.toRadians(t)), cy - ry * Math.sin(Math.toRadians(t))}); any = true; }
            }
            if (!any || pts.size() < 2) continue;
            side.moveTo(pts.get(0)[0], pts.get(0)[1]);
            for (double[] q : pts) side.lineTo(q[0], q[1]);
            for (int k = pts.size() - 1; k >= 0; k--) side.lineTo(pts.get(k)[0], pts.get(k)[1] + th);
            side.closePath();
            Color c = pointColor(g, s, (int) sl[0]);
            cv.fill(side, shade(c, 0.7));
        }
        // 上面
        for (double[] sl : slices) {
            Arc2D top = new Arc2D.Double(cx - rx, cy - ry, 2 * rx, 2 * ry, sl[1], sl[2], Arc2D.PIE);
            cv.fill(top, pointColor(g, s, (int) sl[0]));
            dp.strokeShape(cv, top, edge);
        }
    }

    /**
     * 遠近のある 3D の円グラフ (rAngAx が 0): 半径 0.5 の円盤 (厚さ = 半径 × 0.22 × depthPercent %) を rotX 傾けて遠近をかけ、
     * 上面と下面の縁がプロットに収まる大きさにする。手前の側面、上面の扇の順に描く。
     */
    private void paintPie3DPersp(PdfCanvas cv, Group g, Series s, double total, Rectangle2D box) throws IOException {
        double h = 0.5 * 0.22 * depthPct / 100;
        Proj pj = new Proj(rotX, 0, perspective, h, 1);
        int N = 144;
        double fw = box.getWidth() * 0.9, fh = box.getHeight() * 0.92;
        double eyePt = box.getHeight() / 2 / Math.tan(Math.toRadians(perspective / 4));
        for (int it = 0; it < 30; it++) {
            pj.s = 1; pj.ox = 0; pj.oy = 0;
            double x0 = 1e9, y0 = 1e9, x1 = -1e9, y1 = -1e9;
            for (int k = 0; k < N; k++) {
                double t = 2 * Math.PI * k / N;
                for (double yy : new double[] {0, h}) {
                    double[] q = pj.p(0.5 + 0.5 * Math.cos(t), yy, 0.5 + 0.5 * Math.sin(t));
                    x0 = Math.min(x0, q[0]); x1 = Math.max(x1, q[0]); y0 = Math.min(y0, q[1]); y1 = Math.max(y1, q[1]);
                }
            }
            double sc = Math.min(fw / (x1 - x0), fh / (y1 - y0));
            pj.s = sc; pj.ox = box.getCenterX() - (x0 + x1) / 2 * sc; pj.oy = box.getCenterY() - (y0 + y1) / 2 * sc;
            pj.eye = eyePt / sc;
        }
        final Proj P3 = pj;
        java.util.function.BiFunction<Double, Double, double[]> at = (deg, yy) -> {
            double t = Math.toRadians(deg);
            return P3.p(0.5 + 0.5 * Math.cos(t), yy, 0.5 + 0.5 * Math.sin(t));
        };
        double start = 90 - g.firstSliceAng - rotY;
        List<double[]> slices = new ArrayList<>();
        for (int i = 0; i < s.vals.size(); i++) {
            Double v = s.vals.get(i);
            if (v == null || v <= 0) continue;
            double ext = -v / total * 360;
            slices.add(new double[] {i, start, ext});
            start += ext;
        }
        // 側面 (手前の半分: 角度 180〜360 度)
        for (double[] sl : slices) {
            List<double[]> top = new ArrayList<>(), bot = new ArrayList<>();
            int steps = 48;
            for (int k = 0; k <= steps; k++) {
                double t = sl[1] + sl[2] * k / steps, tn = norm(t);
                if (tn >= 180 && tn <= 360) { top.add(at.apply(t, h)); bot.add(at.apply(t, 0.0)); }
            }
            if (top.size() < 2) continue;
            Color base = pointColor(g, s, (int) sl[0]);
            Path2D side = new Path2D.Double();
            side.moveTo(top.get(0)[0], top.get(0)[1]);
            for (double[] q : top) side.lineTo(q[0], q[1]);
            for (int k = bot.size() - 1; k >= 0; k--) side.lineTo(bot.get(k)[0], bot.get(k)[1]);
            side.closePath();
            cv.fill(side, shade(base, 0.7));
        }
        DrawingPainter.LineStyle edge = defaultLine(dp.bw ? Color.BLACK : Color.WHITE, 0.75);
        double[] c = P3.p(0.5, h, 0.5);
        for (double[] sl : slices) {
            Path2D top = new Path2D.Double();
            top.moveTo(c[0], c[1]);
            int steps = Math.max(4, (int) Math.ceil(Math.abs(sl[2]) / 2));
            for (int k = 0; k <= steps; k++) { double[] q = at.apply(sl[1] + sl[2] * k / steps, h); top.lineTo(q[0], q[1]); }
            top.closePath();
            cv.fill(top, pointColor(g, s, (int) sl[0]));
            dp.strokeShape(cv, top, edge);
        }
    }

    /** 角度を 0〜360 度にする。 */
    private static double norm(double a) {
        double r = a % 360;
        return r < 0 ? r + 360 : r;
    }

    /** 円グラフ・ドーナツグラフを描く。 */
    private void paintPie(PdfCanvas cv, Group g, Rectangle2D box) throws IOException {
        if (g.series.isEmpty()) return;
        if (g.ofPie != null) { paintOfPie(cv, g, box); return; }
        if (g.is3D) { paintPie3D(cv, g, box); return; }
        Series s = g.series.get(0);
        double total = 0;
        for (Double v : s.vals) if (v != null && v > 0) total += v;
        if (total <= 0) return;
        // 旧来のスタイルは短い辺の 0.97 倍 (exp75 の実測)、それ以外は短い辺 − 7pt (exp81 の実測: 短い辺 149pt で 140pt、85pt で 77pt)
        double side = Math.min(box.getWidth(), box.getHeight());
        double d = legacy ? side * LEGACY_PIE : Math.max(side * 0.5, side - PIE_INSET);
        // データラベルを最適な位置・外側に置く円は、ラベルの場所を空けて小さくする (Excel の印刷の実測 2 件: 0.84 倍)
        if (outsideLabels(g, s)) d *= PIE_LABEL_SHRINK;
        // 切り離し (explosion %): 最も離れる扇が元の円に収まるよう、直径を 1 / (1 + 最大の割合) にする
        double emax = 0;
        for (int i = 0; i < s.vals.size(); i++) emax = Math.max(emax, explosion(s, i));
        d /= 1 + emax / 100;
        double cx0 = box.getCenterX(), cy0 = box.getCenterY();
        dumpPlot = new Rectangle2D.Double(cx0 - d / 2, cy0 - d / 2, d, d);
        double start = 90 - g.firstSliceAng; // Java の角度は 3 時の方向から反時計回り
        double hole = g.type.equals("doughnut") ? g.holeSize / 100.0 : 0;
        DrawingPainter.LineStyle edge = s.spPr != null && Xml.child(s.spPr, "ln") != null ? dp.line(Xml.child(s.spPr, "ln"), null, null)
                : (legacy && !dp.bw ? null : defaultLine(dp.bw ? Color.BLACK : Color.WHITE, 0.75));
        for (int i = 0; i < s.vals.size(); i++) {
            Double v = s.vals.get(i);
            if (v == null || v <= 0) continue;
            double ext = -v / total * 360;
            double em = Math.toRadians(start + ext / 2), off = d / 2 * explosion(s, i) / 100;
            double cx = cx0 + Math.cos(em) * off, cy = cy0 - Math.sin(em) * off;
            Rectangle2D circle = new Rectangle2D.Double(cx - d / 2, cy - d / 2, d, d);
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
            if (dp.bw) dp.fillShape(cv, slice, bwFill(c), null); else cv.fill(slice, c);
            if (edge != null) dp.strokeShape(cv, slice, edge);
            double mid = Math.toRadians(start + ext / 2);
            double lr = d / 2 * (hole > 0 ? (1 + hole) / 2 : 0.6);
            pieCx = cx; pieCy = cy; pieR = d / 2; pieMid = mid;
            paintDataLabel(cv, g, s, i, v, new double[] {cx + Math.cos(mid) * lr, cy - Math.sin(mid) * lr}, null, v / total);
            pieR = -1;
            start += ext;
        }
    }

    /**
     * 縦の値の軸のラベルの右端とプロットの間 (pt)。書式のあるグラフは文字の大きさ × 0.91 で、ラベルの左端はグラフの内側の左端の 0.96pt 外
     * (Excel の印刷 exp81: 9pt で 8.2pt、12pt で 10.9pt)。旧来のスタイルは 4pt で、ラベルの左に 2pt の余白。
     */
    private double valLabelGap(TStyle vt) {
        return !modern() ? 4 : vt.size * VAL_GAP_EM;
    }

    static final double VAL_GAP_EM = 0.91;
    /** 書式のあるグラフの値のラベルは、グラフの内側の左端 (外から 7pt) より 0.96pt 外から始まる (Excel の印刷 exp81)。 */
    static final double VAL_LABEL_OUT = 0.96;
    /** 項目軸がプロットの途中で交わるときの、プロットの下の場所 (文字の大きさに対する比。Excel の印刷 exp81)。 */
    static final double NEG_BOTTOM = 0.53;
    static final double LINE_LABEL_GAP = 8.2;
    static final double INSIDE_LEFT = 3.5;
    static final double TREND_LBL_DX = 9.4, TREND_LBL_DY = 2.75;
    static final double SEC_LABEL_EXTRA = 0.8, SEC_RIGHT_CUT = 3.1;

    /**
     * 書式のあるグラフの下の項目名を、従来の位置から下げる量 (pt)。Excel の印刷 exp81: プロットの下端から基線までが
     * 欧文 9pt で 14.6pt、12pt で 19.2pt (従来は 13.5・16.8)、和文 9pt で 16.6pt。
     */
    private double catLabelDrop(TStyle t, List<String> cats) {
        if (legacy) return hasCjk(cats) ? 1.5 : 0; // 旧来のスタイルの和文の項目名は 1.5pt 下 (Excel の印刷 exp90)
        if (!modern()) return 0;
        double d = 0.42 * t.size - 2.7;
        return hasCjk(cats) ? d + 0.22 * t.size : d;
    }

    /** 書式のあるグラフの下の項目名の場所に足す量 (pt)。欧文は +1.3、和文は −1.5 (Excel の印刷 exp81)。 */
    private double catBottomExtra(TStyle t, List<String> cats) {
        if (legacy) return hasCjk(cats) ? -4 : 0; // 旧来のスタイルの和文の項目名の場所は 4pt 少なく (exp90)
        if (!modern()) return 0;
        return hasCjk(cats) ? -1.5 : 1.3;
    }

    private static boolean hasCjk(List<String> texts) {
        for (String x : texts) if (x != null && x.codePoints().anyMatch(FontManager::isCjk)) return true;
        return false;
    }

    /** 横の値の軸の目盛りの間隔の最小 (pt)。旧来のスタイルは 36、書式のあるグラフは 30 (Excel の印刷 exp81 の散布図: 32.8pt の間隔で目盛りを描く)。 */
    private double hGap() {
        return legacy ? 30 : !modern() ? 36 : H_GAP; // 旧来のスタイルは 30 (Excel の印刷 exp90 の横棒)
    }

    static final double H_GAP = 30;
    /** 散布図の x の目盛りを決める長さ = グラフの幅 − この値 (書式のあるグラフ。Excel の印刷 exp71・72・81)。 */
    static final double H_LEN_CUT = 20;

    /** 要素の切り離しの割合 (%): 要素の dPt の explosion、なければ系列の explosion。 */
    private static double explosion(Series s, int i) {
        if (s.ser == null) return 0;
        for (Element dpt : Xml.children(s.ser, "dPt"))
            if ((int) Xml.doubleAttr(Xml.child(dpt, "idx"), "val", -1) == i && Xml.child(dpt, "explosion") != null)
                return Math.max(0, Xml.doubleAttr(Xml.child(dpt, "explosion"), "val", 0));
        return Math.max(0, Xml.doubleAttr(Xml.child(s.ser, "explosion"), "val", 0));
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
        TStyle vt = style(Xml.child(va.el, "txPr"));
        Axis ca = null;
        for (Axis a : axes) if (a.kind.equals("catAx")) ca = a;
        TStyle ct = style(ca == null ? null : Xml.child(ca.el, "txPr"));
        // 項目名の分を空けた半径
        double rad = Math.min(box.getWidth(), box.getHeight()) / 2 - textLineH(ct, g.series.isEmpty() ? List.of() : g.series.get(0).cats) - (legacy ? RADAR_LABEL_GAP : 1.1);
        rad = Math.max(rad, Math.min(box.getWidth(), box.getHeight()) / 4);
        autoScale(va, 0, max, rad, lineH(vt) * 0.6);
        // レーダーは上端の余裕を取らない (Excel: 最大 6 で 0〜6)
        if (va.max == null && !va.logScale && va.hi - va.unit >= max - 1e-9 && va.hi - va.unit > va.lo) va.hi -= va.unit;
        double cx = box.getCenterX(), cy = box.getCenterY();
        DrawingPainter.LineStyle grid = gridDefault();
        // 中心からの線 (項目ごと)
        for (int i = 0; i < n; i++) {
            double a = Math.PI / 2 - 2 * Math.PI * i / n;
            dp.strokeShape(cv, new Line2D.Double(cx, cy, cx + Math.cos(a) * rad, cy - Math.sin(a) * rad), grid);
        }
        for (double v = va.lo; v <= va.hi + 1e-9; v = step(va, v)) {
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
                Double v = i % n < s.vals.size() ? s.vals.get(i % n) : null;   // 値の少ない系列 (rand30 rd22)
                double rr = ((v == null ? 0 : v) - va.lo) / (va.hi - va.lo) * rad;
                double a = Math.PI / 2 - 2 * Math.PI * i / n;
                double x = cx + Math.cos(a) * rr, y = cy - Math.sin(a) * rr;
                if (i == 0) p.moveTo(x, y); else p.lineTo(x, y);
            }
            if (g.radarStyle.equals("filled")) cv.fill(p, s.color);
            else dp.strokeShape(cv, p, seriesLine(s));
            if (g.radarStyle.equals("marker") && g.markers) {
                for (int i = 0; i < n && i < s.vals.size(); i++) {
                    Double v = s.vals.get(i);
                    if (v == null) continue;
                    double rr = (v - va.lo) / (va.hi - va.lo) * rad, a = Math.PI / 2 - 2 * Math.PI * i / n;
                    marker(cv, s, si, cx + Math.cos(a) * rr, cy - Math.sin(a) * rr);
                }
            }
        }
        // 値の目盛ラベル (上向きの線の左)
        if (!va.deleted && !"none".equals(va.tickLblPos)) {
            for (double v = va.lo; v <= va.hi + 1e-9; v = step(va, v)) {
                double rr = (v - va.lo) / (va.hi - va.lo) * rad;
                text(cv, axisLabel(va, v, g.series.isEmpty() ? "General" : g.series.get(0).fmt), cx - vt.size * (legacy ? LEGACY_RADAR_VAL_GAP : RADAR_VAL_GAP), cy - rr, vt, 1, 0);
            }
        }
        // 項目名 (線の先の外側。向きに合わせて揃える)
        if (ca == null || (!ca.deleted && !"none".equals(ca.tickLblPos))) {
            List<String> cats = g.series.isEmpty() ? List.of() : g.series.get(0).cats;
            for (int i = 0; i < n; i++) {
                double a = Math.PI / 2 - 2 * Math.PI * i / n, co = Math.cos(a), si = Math.sin(a);
                String lab = i < cats.size() ? cats.get(i) : String.valueOf(i + 1);
                // 真上・真下だけ上下に寄せ、斜めの項目名は上下の中央を線の先に合わせる (Excel の印刷 exp81)
                int ha = co > 0.3 ? -1 : co < -0.3 ? 1 : 0, vaa = si > 0.95 ? 1 : si < -0.95 ? -1 : 0; // 7 項目の下 (sin -0.90) は中央 (exp75)
                text(cv, lab, cx + co * (rad + 3), cy - si * (rad + 3), ct, ha, vaa);
            }
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
        // データ要素ごとの指定 (手動で動かした位置 manualLayout の x, y はグラフの幅・高さに対する移動量)
        double mx = Double.NaN, my = Double.NaN;
        for (Element d : Xml.children(dl, "dLbl")) {
            if (Xml.intAttr(Xml.child(d, "idx"), "val", -1) == i) {
                if (Xml.boolAttr(Xml.child(d, "delete"), "val", false)) return;
                if (Xml.child(d, "showVal") != null) showVal = Xml.boolAttr(Xml.child(d, "showVal"), "val", false);
                Element ml = Xml.path(d, "layout", "manualLayout");
                if (ml != null && Xml.child(ml, "x") != null && !"edge".equals(Xml.attr(Xml.child(ml, "xMode"), "val", "factor"))) {
                    mx = Xml.doubleAttr(Xml.child(ml, "x"), "val", 0);
                    my = Xml.doubleAttr(Xml.child(ml, "y"), "val", 0);
                }
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
                if (!Double.isNaN(mx) && pieR > 0 && chartRect != null) {
                    // 動かしたラベル: 扇の外側 (outEnd) の位置から移動し、引き出し線で扇の縁とつなぐ
                    double hw = textWidth(line, t) / 2, hh = lineH(t) / 2;
                    double co = Math.cos(pieMid), si = -Math.sin(pieMid);
                    double bx = pieCx + co * (pieR + hw + 2), by = pieCy + si * (pieR + hh + 2);
                    double lx = bx + mx * chartRect.getWidth(), ly = by + my * chartRect.getHeight();
                    text(cv, line, lx, ly, t, 0, 0);
                    boolean leader = Xml.boolAttr(Xml.child(dl, "showLeaderLines"), "val", false);
                    double ex = pieCx + co * pieR, ey = pieCy + si * pieR;
                    // ラベルの矩形の、扇に近い角・辺の点
                    double tx = Math.max(lx - hw, Math.min(lx + hw, ex)), ty = Math.max(ly - hh, Math.min(ly + hh, ey));
                    if (leader && Math.hypot(tx - ex, ty - ey) > 3) {
                        Element ll = Xml.child(Xml.child(dl, "leaderLines"), "spPr");
                        DrawingPainter.LineStyle ls = ll != null && Xml.child(ll, "ln") != null ? dp.line(Xml.child(ll, "ln"), null, null)
                                : defaultLine(legacy ? Color.BLACK : new Color(0xA6A6A6), 0.75);
                        if (ls != null) {
                            // ラベルが扇の縁より横にあれば、ラベルの中ほどへ短い横線で折れて入る (Excel 2013 以降)
                            Path2D pl = new Path2D.Double();
                            pl.moveTo(ex, ey);
                            if (lx - hw > ex + LEADER_ELBOW) { pl.lineTo(lx - hw - 1 - LEADER_ELBOW, ly); pl.lineTo(lx - hw - 1, ly); }
                            else if (lx + hw < ex - LEADER_ELBOW) { pl.lineTo(lx + hw + 1 + LEADER_ELBOW, ly); pl.lineTo(lx + hw + 1, ly); }
                            else pl.lineTo(tx, ty);
                            dp.strokeShape(cv, pl, ls);
                        }
                    }
                    continue;
                }
                text(cv, line, x, y, t, 0, 0);
                y += lineH(t);
            } else {
                switch (pos) {
                    case "t" -> text(cv, line, x, y - 4, t, 0, 1);
                    case "b" -> text(cv, line, x, y + 4, t, 0, -1);
                    // 左右は点から 5pt (書式のあるグラフは 8.2pt。Excel の印刷 exp81 のマーカー付き折れ線)
                    case "l" -> text(cv, line, x - (modern() ? LINE_LABEL_GAP : 5), y, t, 1, 0);
                    case "ctr" -> text(cv, line, x, y, t, 0, 0);
                    default -> text(cv, line, x + (modern() ? LINE_LABEL_GAP : 5), y, t, -1, 0);
                }
            }
        }
    }

    /** データラベルが外側に出うる円の大きさの倍率。 */
    static final double PIE_LABEL_SHRINK = 0.84;

    /** 円グラフのデータラベルが表示され、位置が最適 (bestFit)・外側 (outEnd) か。 */
    private static boolean outsideLabels(Group g, Series s) {
        Element dl = s.dLbls != null ? s.dLbls : g.dLbls;
        if (dl == null || Xml.boolAttr(Xml.child(dl, "delete"), "val", false)) return false;
        boolean shown = false;
        for (String k : new String[] {"showVal", "showPercent", "showCatName", "showSerName"})
            if (Xml.boolAttr(Xml.child(dl, k), "val", false)) shown = true;
        String pos = Xml.attr(Xml.child(dl, "dLblPos"), "val", "bestFit");
        return shown && (pos.equals("bestFit") || pos.equals("outEnd"));
    }

    /** 書式のないグラフの縦の値の軸の目盛りの最小の間隔 (文字の大きさの倍。Excel の印刷: 10pt で 12.9pt 間隔の 1 刻み)。 */
    static final double LEGACY_GAP = 1.35; // 1.25 では小さいグラフで細かすぎ (exp90)、1.5 では exp48 が粗すぎ

    /** 引き出し線の、ラベルの手前の横の部分の長さ (pt)。 */
    private static final double LEADER_ELBOW = 5;

    /** 数値を小数 2 桁の文字列にする (デバッグ用)。 */
    static String fmt(double d) {
        return String.format(Locale.ROOT, "%.2f", d);
    }
}
