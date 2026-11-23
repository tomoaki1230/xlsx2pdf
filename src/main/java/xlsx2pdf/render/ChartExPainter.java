package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.Arc2D;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.w3c.dom.Element;

/**
 * Excel 2016 以降のグラフ (chartEx: ウォーターフォール、じょうご、ヒストグラム、パレート図、箱ひげ図、ツリーマップ、サンバースト)。
 * データは cx:f の名前 (_xlchart.v1.N) が指すセル範囲から読む。書式は 2013 以降の既定 (文字 #595959、目盛線 #D9D9D9)。
 */
final class ChartExPainter {
    private final RenderContext ctx;
    private final PackagePart part;
    private static final Color TEXT = new Color(0x595959), GRID = new Color(0xD9D9D9);
    private static final float SIZE = 9;

    ChartExPainter(RenderContext ctx, PackagePart part) {
        this.ctx = ctx;
        this.part = part;
    }

    /** 1 つのデータ (項目の階層と値)。 */
    private static final class Data {
        /** 項目の階層 (外側の階層から順。各要素は点ごとの文字列)。 */
        List<List<String>> levels = new ArrayList<>();
        List<Double> vals = new ArrayList<>();
    }

    /** グラフを矩形 r に描く。 */
    void paint(PdfCanvas cv, Rectangle2D r) throws IOException {
        Element root;
        try (InputStream in = part.getInputStream()) {
            root = Xml.parse(in);
        } catch (Exception e) {
            Warnings.add("グラフの定義を読めませんでした", e);
            return;
        }
        Map<String, Data> data = new LinkedHashMap<>();
        for (Element d : Xml.children(Xml.child(root, "chartData"), "data")) data.put(d.getAttribute("id"), readData(d));
        Element chart = Xml.child(root, "chart");
        if (chart == null) return;
        cv.fillRect(r.getX(), r.getY(), r.getWidth(), r.getHeight(), Color.WHITE);
        cv.save();
        cv.clipRect(r.getX(), r.getY(), r.getWidth(), r.getHeight());
        double pad = 7;
        Rectangle2D inner = new Rectangle2D.Double(r.getX() + pad, r.getY() + pad, r.getWidth() - 2 * pad, r.getHeight() - 2 * pad);
        double top = inner.getY();
        Element title = Xml.child(chart, "title");
        if (title != null) {
            String t = richText(title);
            if (t.isEmpty()) t = "グラフ タイトル";
            text(cv, t, inner.getCenterX(), top, 14, TEXT, 0, -1);
            top += lineH(14) + 4;
        }
        Element pa = Xml.child(chart, "plotArea");
        List<Element> series = Xml.children(Xml.child(pa, "plotAreaRegion"), "series");
        if (series.isEmpty()) { cv.restore(); return; }
        Element s0 = series.get(0);
        String layout = s0.getAttribute("layoutId");
        Data d = data.get(Xml.attr(Xml.child(s0, "dataId"), "val", "0"));
        if (d == null) { cv.restore(); return; }
        Element legend = Xml.child(chart, "legend");
        Rectangle2D box = new Rectangle2D.Double(inner.getX(), top, inner.getWidth(), inner.getMaxY() - top);
        try {
            switch (layout) {
                case "waterfall" -> {
                    if (legend != null) box = legend(cv, box, List.of(new Object[] {"増加", accent(0)}, new Object[] {"減少", accent(1)}, new Object[] {"合計", accent(2)}));
                    waterfall(cv, s0, d, box);
                }
                case "funnel" -> funnel(cv, d, box);
                case "clusteredColumn" -> {
                    boolean pareto = series.size() > 1 && "paretoLine".equals(series.get(1).getAttribute("layoutId"));
                    columns(cv, s0, d, box, pareto);
                }
                case "boxWhisker" -> boxWhisker(cv, s0, d, box);
                case "treemap" -> {
                    if (legend != null && !d.levels.isEmpty()) box = legend(cv, box, topNames(d));
                    treemap(cv, d, box);
                }
                case "sunburst" -> sunburst(cv, d, box);
                default -> Warnings.add("このグラフの種類 (" + layout + ") は描けません", null);
            }
        } finally {
            cv.restore();
        }
        // 外枠は ChartPainter と同じく左上に寄せる (Excel の印刷の実測)
        cv.stroke(new Rectangle2D.Double(r.getX() + ChartPainter.FRAME_DX, r.getY() + ChartPainter.FRAME_DY, r.getWidth(), r.getHeight()),
                new PdfCanvas.Stroke(0.75f), GRID);
    }

    // ------------------------------------------------------------------ データ

    private Data readData(Element d) {
        Data out = new Data();
        for (Element dim : Xml.children(d)) {
            String n = dim.getLocalName();
            if (!n.equals("strDim") && !n.equals("numDim")) continue;
            String f = Xml.text(Xml.child(dim, "f")).trim();
            List<List<String>> cols = new ArrayList<>();
            if (!f.isEmpty()) cols = range(f);
            if (cols.isEmpty()) {
                // 保存された値 (cx:lvl/cx:pt)
                for (Element lvl : Xml.children(dim, "lvl")) {
                    List<String> l = new ArrayList<>();
                    int cnt = Xml.intAttr(lvl, "ptCount", 0);
                    for (int i = 0; i < cnt; i++) l.add(null);
                    for (Element pt : Xml.children(lvl, "pt")) {
                        int i = Xml.intAttr(pt, "idx", 0);
                        if (i >= 0 && i < l.size()) l.set(i, Xml.text(pt));
                    }
                    cols.add(0, l);
                }
            }
            if (n.equals("numDim")) {
                if (!cols.isEmpty()) for (String v : cols.get(cols.size() - 1)) out.vals.add(parse(v));
            } else out.levels = cols;
        }
        return out;
    }

    private static Double parse(String v) {
        if (v == null || v.isBlank()) return null;
        try { return Double.parseDouble(v.trim()); } catch (NumberFormatException e) { return null; }
    }

    /** 名前 (_xlchart.v1.N) またはセル範囲の値。列ごとの文字列の並び。 */
    private List<List<String>> range(String f) {
        List<List<String>> out = new ArrayList<>();
        String ref = f;
        var name = ctx.wb.getName(f);
        if (name != null) ref = name.getRefersToFormula();
        AreaReference ar;
        try {
            ar = new AreaReference(ref, SpreadsheetVersion.EXCEL2007);
        } catch (RuntimeException e) {
            return out;
        }
        CellReference a = ar.getFirstCell(), b = ar.getLastCell();
        Sheet sh = a.getSheetName() == null ? null : ctx.wb.getSheet(a.getSheetName());
        if (sh == null) return out;
        long total = 0;
        for (int c = a.getCol(); c <= b.getCol(); c++) {
            // 行全体の参照 (16384 列 × 10 万行) で終わらなくならないよう、全体で ChartPainter と同じ件数までにする
            if (total >= ChartPainter.MAX_POINTS) break;
            List<String> col = new ArrayList<>();
            for (int rr = a.getRow(); rr <= b.getRow() && rr - a.getRow() < 100_000; rr++) {
                Row row = sh.getRow(rr);
                Cell cell = row == null ? null : row.getCell(c);
                String v = null;
                if (cell != null) {
                    CellType t = cell.getCellType();
                    if (t == CellType.FORMULA) t = cell.getCachedFormulaResultType();
                    if (t == CellType.NUMERIC) v = String.valueOf(cell.getNumericCellValue());
                    else if (t == CellType.STRING) v = cell.getStringCellValue();
                }
                col.add(v);
            }
            total += col.size();
            out.add(col);
        }
        return out;
    }

    // ------------------------------------------------------------------ 共通

    private Color accent(int i) {
        String[] acc = {"accent1", "accent2", "accent3", "accent4", "accent5", "accent6"};
        Color c = ctx.theme.scheme.get(acc[i % 6]);
        return c == null ? Color.GRAY : c;
    }

    private FontManager.Spec spec() {
        return ctx.fonts.spec(ctx.theme.fontRef("+mn-lt", false), ctx.theme.fontRef("+mn-ea", true), false, false);
    }

    /**
     * 項目ラベルの行の場所 (pt)。和文を含めば行の高さ + 6、欧文だけなら欧文の書体の行 + 5
     * (9pt の欧文の項目で 16.6pt)。
     */
    private double catSpace(Data d) {
        for (List<String> l : d.levels) for (String x : l) if (x != null && x.codePoints().anyMatch(FontManager::isCjk)) return lineH(SIZE) + 6;
        FontManager.Spec sp = spec();
        return (sp.face.ascent + sp.face.descent) * SIZE + 5;
    }

    private double lineH(float size) {
        FontManager.Spec sp = spec();
        return (sp.ascent + sp.descent) * size;
    }

    private double width(String s, float size) { return spec().width(s, size, ctx.fonts); }

    /** 文字を基準位置に揃えて描く (halign: -1 左 0 中央 1 右、valign: -1 上 0 中央 1 下)。 */
    private void text(PdfCanvas cv, String s, double x, double y, float size, Color c, int ha, int va) throws IOException {
        FontManager.Spec sp = spec();
        double w = sp.width(s, size, ctx.fonts), asc = sp.ascent * size, desc = sp.descent * size;
        double xx = ha < 0 ? x : ha == 0 ? x - w / 2 : x - w;
        double base = va < 0 ? y + asc : va == 0 ? y + (asc - desc) / 2 : y - desc;
        cv.text(s, xx, base, sp, size, c, 1, 0);
    }

    private static String richText(Element title) {
        StringBuilder sb = new StringBuilder();
        Element rich = Xml.path(title, "tx", "txData");
        if (rich == null) rich = Xml.path(title, "tx", "rich");
        if (rich == null) return "";
        for (Element p : Xml.children(rich)) {
            if (p.getLocalName().equals("v")) sb.append(Xml.text(p));
            for (Element rr : Xml.children(p, "r")) sb.append(Xml.text(Xml.child(rr, "t")));
        }
        return sb.toString();
    }

    private String fmt(double v) { return ctx.formatter.formatNumber(Math.abs(v) < 1e-12 ? 0 : v, "General"); }

    /** 上の凡例を描き、残りの矩形を返す。項目は {名前, 色}。 */
    private Rectangle2D legend(PdfCanvas cv, Rectangle2D box, List<Object[]> items) throws IOException {
        double total = 0;
        for (Object[] it : items) total += 7 + 3 + width((String) it[0], SIZE) + 10;
        total -= 10;
        // ウォーターフォールの凡例: 印は 7pt の四角、行の中心は題名の下 + 5.5pt、プロットは凡例の行の下 + 16.5pt
        double x = box.getCenterX() - total / 2, y = box.getY() + LEGEND_DROP;
        double lh = lineH(SIZE);
        for (Object[] it : items) {
            cv.fillRect(x, y + lh / 2 - 3.5, 7, 7, (Color) it[1]);
            x += 9;
            text(cv, (String) it[0], x, y + lh / 2, SIZE, TEXT, -1, 0);
            x += width((String) it[0], SIZE) + 10;
        }
        double used = LEGEND_DROP + lh + LEGEND_GAP;
        return new Rectangle2D.Double(box.getX(), box.getY() + used, box.getWidth(), box.getHeight() - used);
    }

    /** 値のラベルの幅に足すプロットの左の余白、右の余白、ラベルとプロットの間 (pt)。 */
    static final double PLOT_LEFT = 4.1, PLOT_RIGHT = -0.3, VAL_LABEL_GAP = 5.9;

    /** 凡例の行を下げる量と、凡例の行とプロットの間 (pt)。 */
    static final double LEGEND_DROP = 5.5, LEGEND_GAP = 11;

    /** 値の軸 (縦) と目盛線を描き、プロットの矩形を返す。lo/hi/unit は ChartPainter.Axis に入れる。 */
    private Rectangle2D valueAxis(PdfCanvas cv, Rectangle2D box, ChartPainter.Axis ax, double dmin, double dmax, double bottomSpace,
                                  boolean percentRight, ChartPainter.Axis ax2) throws IOException {
        ax.maxIntervals = 10; // Excel 2016 以降のグラフは目盛りの区間を 10 まで (0〜55 の 5 刻みでなく 0〜60 の 10 刻み)
        ChartPainter.autoScale(ax, dmin, dmax, box.getHeight() - bottomSpace, lineH(SIZE));
        double lw = 0;
        for (double v = ax.lo; v <= ax.hi + ax.unit * 1e-6; v += ax.unit) lw = Math.max(lw, width(fmt(v), SIZE));
        double rw = percentRight ? width("100%", SIZE) + 6 : 0;
        // プロットの左右の余白: 右に % の軸がないものは左 4.1pt・右 -0.3pt
        // パレート図 (右に % の軸): 左 9.8pt・右 4pt・上 5pt
        double pl = percentRight ? 9.8 : PLOT_LEFT, pr = percentRight ? 4 : PLOT_RIGHT;
        double pt = percentRight ? 5 : 4;
        Rectangle2D plot = new Rectangle2D.Double(box.getX() + lw + pl, box.getY() + pt, box.getWidth() - lw - pl - pr - rw, box.getHeight() - bottomSpace - pt);
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.75f);
        for (double v = ax.lo; v <= ax.hi + ax.unit * 1e-6; v += ax.unit) {
            double y = vy(ax, v, plot);
            cv.line(plot.getX(), y, plot.getMaxX(), y, st, GRID);
            text(cv, fmt(v), plot.getX() - VAL_LABEL_GAP, y, SIZE, TEXT, 1, 0);
        }
        if (percentRight && ax2 != null) {
            for (double v = ax2.lo; v <= ax2.hi + 1e-9; v += ax2.unit) {
                text(cv, Math.round(v * 100) + "%", plot.getMaxX() + 4, vy(ax2, v, plot), SIZE, TEXT, -1, 0);
            }
        }
        return plot;
    }

    private static double vy(ChartPainter.Axis a, double v, Rectangle2D plot) {
        return plot.getMaxY() - (v - a.lo) / (a.hi - a.lo) * plot.getHeight();
    }

    /** 項目名を区画の下に描く。 */
    private void catLabels(PdfCanvas cv, List<String> labels, Rectangle2D plot) throws IOException {
        int n = labels.size();
        double slot = plot.getWidth() / Math.max(1, n);
        for (int i = 0; i < n; i++) text(cv, labels.get(i) == null ? "" : labels.get(i), plot.getX() + (i + 0.5) * slot, plot.getMaxY() + 4, SIZE, TEXT, 0, -1);
        cv.line(plot.getX(), plot.getMaxY(), plot.getMaxX(), plot.getMaxY(), new PdfCanvas.Stroke(0.75f), GRID);
    }

    private List<String> cats(Data d, int n) {
        List<String> out = new ArrayList<>();
        List<String> l = d.levels.isEmpty() ? List.of() : d.levels.get(d.levels.size() - 1);
        for (int i = 0; i < n; i++) out.add(i < l.size() && l.get(i) != null ? l.get(i) : String.valueOf(i + 1));
        return out;
    }

    // ------------------------------------------------------------------ 種類ごと

    /** ウォーターフォール: 合計 (subtotals の番号) は 0 から、それ以外は前の累計から伸ばす。 */
    private void waterfall(PdfCanvas cv, Element s, Data d, Rectangle2D box) throws IOException {
        int n = d.vals.size();
        java.util.Set<Integer> totals = new java.util.HashSet<>();
        for (Element i : Xml.children(Xml.path(s, "layoutPr", "subtotals"), "idx")) totals.add(Xml.intAttr(i, "val", -1));
        double[] from = new double[n], to = new double[n];
        double cum = 0, mn = 0, mx = 0;
        for (int i = 0; i < n; i++) {
            double v = d.vals.get(i) == null ? 0 : d.vals.get(i);
            if (totals.contains(i)) { from[i] = 0; to[i] = v; cum = v; }
            else { from[i] = cum; to[i] = cum + v; cum += v; }
            mn = Math.min(mn, Math.min(from[i], to[i]));
            mx = Math.max(mx, Math.max(from[i], to[i]));
        }
        ChartPainter.Axis ax = new ChartPainter.Axis();
        Rectangle2D plot = valueAxis(cv, box, ax, mn, mx, catSpace(d), false, null);
        double gap = catGap(s, 0.5);
        double slot = plot.getWidth() / Math.max(1, n), bw = slot / (1 + gap);
        boolean labels = Xml.child(s, "dataLabels") != null;
        // 棒をつなぐ線 (累計の高さ、棒の右端から次の棒の左端。#D9D9D9 の 0.75pt)
        Element vis = Xml.child(Xml.child(s, "layoutPr"), "visibility");
        if (vis == null || Xml.boolAttr(vis, "connectorLines", true)) {
            for (int i = 0; i + 1 < n; i++) {
                double y = vy(ax, to[i], plot), x0 = plot.getX() + (i + 0.5) * slot + bw / 2, x1 = plot.getX() + (i + 1.5) * slot - bw / 2;
                cv.line(x0, y, x1, y, new PdfCanvas.Stroke(0.75f), new Color(0xD9D9D9));
            }
        }
        for (int i = 0; i < n; i++) {
            double v = d.vals.get(i) == null ? 0 : d.vals.get(i);
            Color c = totals.contains(i) ? accent(2) : v >= 0 ? accent(0) : accent(1);
            double y0 = vy(ax, Math.max(from[i], to[i]), plot), y1 = vy(ax, Math.min(from[i], to[i]), plot);
            double x = plot.getX() + (i + 0.5) * slot - bw / 2;
            cv.fillRect(x, y0, bw, Math.max(0.5, y1 - y0), c);
            if (labels) {
                if (v >= 0 || totals.contains(i)) text(cv, fmt(v), x + bw / 2, y0 - 1, SIZE, TEXT, 0, 1);
                else text(cv, fmt(v), x + bw / 2, y1 + 1, SIZE, TEXT, 0, -1);
            }
        }
        catLabels(cv, cats(d, n), plot);
    }

    /** じょうご: 値に比例した横棒を中央に並べ、値を棒の中に描く。項目名は左。 */
    private void funnel(PdfCanvas cv, Data d, Rectangle2D box) throws IOException {
        int n = d.vals.size();
        double max = 0;
        for (Double v : d.vals) if (v != null) max = Math.max(max, v);
        if (max <= 0) return;
        List<String> cats = cats(d, n);
        double lw = 0;
        for (String c : cats) lw = Math.max(lw, width(c, SIZE));
        // 項目名とプロットの間 4.2pt、上 0.5pt、下 0.1pt
        Rectangle2D plot = new Rectangle2D.Double(box.getX() + lw + 4.16, box.getY() + 0.5, box.getWidth() - lw - 3.96, box.getHeight() - 0.6);
        double slot = plot.getHeight() / n, bh = slot / 1.06;
        for (int i = 0; i < n; i++) {
            double v = d.vals.get(i) == null ? 0 : d.vals.get(i);
            double w = plot.getWidth() * v / max, y = plot.getY() + i * slot + (slot - bh) / 2;
            cv.fillRect(plot.getCenterX() - w / 2, y, w, bh, accent(0));
            text(cv, fmt(v), plot.getCenterX(), y + bh / 2, SIZE, Color.WHITE, 0, 0);
            text(cv, cats.get(i), plot.getX() - 6, y + bh / 2, SIZE, TEXT, 1, 0);
        }
    }

    /** 有効数字 2 桁に丸める。 */
    private static double sig2(double w) {
        if (!(w > 0)) return w;
        double m = Math.pow(10, Math.floor(Math.log10(w)) - 1);
        return Math.round(w / m) * m;
    }

    /** ヒストグラム (自動の区間: スコットの選択) とパレート図 (降順の棒と累積の割合の折れ線)。 */
    private void columns(PdfCanvas cv, Element s, Data d, Rectangle2D box, boolean pareto) throws IOException {
        List<String> labels = new ArrayList<>();
        List<Double> heights = new ArrayList<>();
        Element binning = Xml.path(s, "layoutPr", "binning");
        if (binning != null || (d.levels.isEmpty() && !pareto)) {
            List<Double> v = new ArrayList<>();
            for (Double x : d.vals) if (x != null) v.add(x);
            if (v.isEmpty()) return;
            java.util.Collections.sort(v);
            double mn = v.get(0), mx = v.get(v.size() - 1);
            double mean = 0, sd = 0;
            for (double x : v) mean += x;
            mean /= v.size();
            for (double x : v) sd += (x - mean) * (x - mean);
            sd = v.size() > 1 ? Math.sqrt(sd / (v.size() - 1)) : 1;
            double w = Xml.child(binning, "binSize") != null ? Xml.doubleAttr(Xml.child(binning, "binSize"), "val", 1)
                    : sig2(3.5 * sd / Math.cbrt(v.size())); // Excel は幅を有効数字 2 桁に丸める
            int bins = Xml.child(binning, "binCount") != null ? Xml.intAttr(Xml.child(binning, "binCount"), "val", 5)
                    : (int) Math.max(1, Math.ceil((mx - mn) / Math.max(w, 1e-12)));
            // 区間の数は 1〜MAX_BINS (binSize が極端に小さい・binCount が極端に大きい・0 以下のファイルで止まらないように)
            if (bins < 1 || bins > MAX_BINS) {
                bins = Math.max(1, Math.min(MAX_BINS, bins));
                w = (mx - mn) / bins;
            }
            if (Xml.child(binning, "binCount") != null || !(w > 0)) w = (mx - mn) / bins;
            boolean right = !"l".equals(Xml.attr(binning, "intervalClosed", "r"));
            for (int b = 0; b < bins; b++) {
                double a = mn + b * w, e = mn + (b + 1) * w;
                int cnt = 0;
                for (double x : v) {
                    boolean in = right ? (b == 0 ? x >= a : x > a) && x <= e : x >= a && (b == bins - 1 ? x <= e : x < e);
                    if (in) cnt++;
                }
                heights.add((double) cnt);
                labels.add((right ? (b == 0 ? "[" : "(") : "[") + fmt(round(a)) + ", " + fmt(round(e)) + (right ? "]" : (b == bins - 1 ? "]" : ")")));
            }
        } else {
            List<String> cats = cats(d, d.vals.size());
            List<Integer> idx = new ArrayList<>();
            for (int i = 0; i < d.vals.size(); i++) if (d.vals.get(i) != null) idx.add(i);
            if (pareto) idx.sort((a, b) -> Double.compare(d.vals.get(b), d.vals.get(a)));
            for (int i : idx) { heights.add(d.vals.get(i)); labels.add(cats.get(i)); }
        }
        double mx = 0;
        for (double h : heights) mx = Math.max(mx, h);
        ChartPainter.Axis ax = new ChartPainter.Axis(), ax2 = null;
        if (pareto) { ax2 = new ChartPainter.Axis(); ax2.lo = 0; ax2.hi = 1; ax2.unit = 0.1; }
        Rectangle2D plot = valueAxis(cv, box, ax, 0, mx, catSpace(d), pareto, ax2);
        int n = heights.size();
        double gap = catGap(s, 0);
        double slot = plot.getWidth() / Math.max(1, n), bw = slot / (1 + gap);
        for (int i = 0; i < n; i++) {
            double y0 = vy(ax, heights.get(i), plot);
            cv.fillRect(plot.getX() + (i + 0.5) * slot - bw / 2, y0, bw, plot.getMaxY() - y0, accent(0));
            if (gap == 0) cv.stroke(new Rectangle2D.Double(plot.getX() + (i + 0.5) * slot - bw / 2, y0, bw, plot.getMaxY() - y0), new PdfCanvas.Stroke(0.75f), Color.WHITE);
        }
        if (pareto && ax2 != null) {
            double tot = 0, cum = 0;
            for (double h : heights) tot += h;
            Path2D p = new Path2D.Double();
            for (int i = 0; i < n; i++) {
                cum += heights.get(i);
                double x = plot.getX() + (i + 0.5) * slot, y = vy(ax2, tot > 0 ? cum / tot : 0, plot);
                if (i == 0) p.moveTo(x, y); else p.lineTo(x, y);
            }
            PdfCanvas.Stroke st = new PdfCanvas.Stroke(2.25f);
            st.cap = 1; st.join = 1;
            cv.stroke(p, st, accent(1));
        }
        catLabels(cv, labels, plot);
    }

    /** 項目軸の間隔 (plotArea の axis/catScaling の gapWidth。系列の 2 つ上の要素が plotArea)。 */
    private static double catGap(Element series, double dflt) {
        if (!(series.getParentNode() instanceof Element region) || !(region.getParentNode() instanceof Element pa)) return dflt;
        for (Element a : Xml.children(pa, "axis")) {
            Element cs = Xml.child(a, "catScaling");
            if (cs != null && cs.hasAttribute("gapWidth")) return Xml.doubleAttr(cs, "gapWidth", dflt);
        }
        return dflt;
    }

    private static double round(double v) { return Math.round(v * 1000) / 1000.0; }

    /** 箱ひげ図: 四分位 (既定は排他的)、ひげは 1.5 × 四分位範囲の内側の最小・最大、平均の × 印、外れ値の点。 */
    private void boxWhisker(PdfCanvas cv, Element s, Data d, Rectangle2D box) throws IOException {
        List<Double> v = new ArrayList<>();
        for (Double x : d.vals) if (x != null) v.add(x);
        if (v.size() < 2) return;
        java.util.Collections.sort(v);
        boolean excl = !"inclusive".equals(Xml.attr(Xml.path(s, "layoutPr", "statistics"), "quartileMethod", "exclusive"));
        double q1 = quartile(v, 0.25, excl), q2 = quartile(v, 0.5, excl), q3 = quartile(v, 0.75, excl), iqr = q3 - q1;
        double lo = v.get(v.size() - 1), hi = v.get(0);
        for (double x : v) { if (x >= q1 - 1.5 * iqr) lo = Math.min(lo, x); if (x <= q3 + 1.5 * iqr) hi = Math.max(hi, x); }
        double mean = 0;
        for (double x : v) mean += x;
        mean /= v.size();
        ChartPainter.Axis ax = new ChartPainter.Axis();
        ax.exactHeadroom = true; // 箱ひげ図は最大値が目盛りに重なるときだけ上に 1 目盛り足す (97 で 0〜100)
        Rectangle2D plot = valueAxis(cv, box, ax, Math.min(0, v.get(0)), v.get(v.size() - 1), catSpace(d), false, null);
        double cx = plot.getCenterX(), bw = plot.getWidth() / 2;
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.75f);
        // ひげ・箱の枠・中央値・平均の印は系列の色を 0.8 倍に暗くした色、ひげの横線は箱の幅の 0.3 倍
        Color f = accent(0), line = new Color((int) (f.getRed() * 0.8), (int) (f.getGreen() * 0.8), (int) (f.getBlue() * 0.8));
        double cap = bw * 0.15;
        cv.line(cx, vy(ax, hi, plot), cx, vy(ax, q3, plot), st, line);
        cv.line(cx, vy(ax, q1, plot), cx, vy(ax, lo, plot), st, line);
        cv.line(cx - cap, vy(ax, hi, plot), cx + cap, vy(ax, hi, plot), st, line);
        cv.line(cx - cap, vy(ax, lo, plot), cx + cap, vy(ax, lo, plot), st, line);
        Rectangle2D b = new Rectangle2D.Double(cx - bw / 2, vy(ax, q3, plot), bw, vy(ax, q1, plot) - vy(ax, q3, plot));
        cv.fillRect(b.getX(), b.getY(), b.getWidth(), b.getHeight(), f);
        cv.stroke(b, st, line);
        cv.line(b.getX(), vy(ax, q2, plot), b.getMaxX(), vy(ax, q2, plot), st, line);
        double my = vy(ax, mean, plot);
        cv.line(cx - 3, my - 3, cx + 3, my + 3, st, line);
        cv.line(cx - 3, my + 3, cx + 3, my - 3, st, line);
        for (double x : v) if (x < lo || x > hi) cv.stroke(new java.awt.geom.Ellipse2D.Double(cx - 2.5, vy(ax, x, plot) - 2.5, 5, 5), st, accent(0));
        catLabels(cv, List.of("1"), plot);
    }

    /** 四分位数 (排他的は QUARTILE.EXC、包括的は QUARTILE.INC)。 */
    private static double quartile(List<Double> v, double p, boolean exclusive) {
        int n = v.size();
        double k = exclusive ? p * (n + 1) - 1 : p * (n - 1);
        k = Math.max(0, Math.min(n - 1, k));
        int i = (int) Math.floor(k);
        double f = k - i;
        return i + 1 < n ? v.get(i) + f * (v.get(i + 1) - v.get(i)) : v.get(i);
    }

    /** 最上位の階層の名前と色 (凡例・ツリーマップの色)。 */
    private List<Object[]> topNames(Data d) {
        List<Object[]> out = new ArrayList<>();
        List<String> seen = new ArrayList<>();
        String last = null;
        for (String s : d.levels.get(0)) {
            if (s == null) s = last;
            last = s;
            if (s != null && !seen.contains(s)) { seen.add(s); out.add(new Object[] {s, accent(seen.size() - 1)}); }
        }
        return out;
    }

    /** ツリーマップ: 最上位の階層ごとの面積で分け、その中を葉の値で分ける (縦横の長い方を分ける並べ方)。 */
    private void treemap(PdfCanvas cv, Data d, Rectangle2D box) throws IOException {
        Map<String, List<Integer>> groups = groups(d);
        List<String> names = new ArrayList<>(groups.keySet());
        double[] sizes = new double[names.size()];
        for (int g = 0; g < names.size(); g++) for (int i : groups.get(names.get(g))) sizes[g] += val(d, i);
        List<Rectangle2D> rs = slice(sizes, box);
        List<String> leaf = d.levels.isEmpty() ? List.of() : d.levels.get(d.levels.size() - 1);
        for (int g = 0; g < names.size(); g++) {
            List<Integer> items = groups.get(names.get(g));
            double[] sz = new double[items.size()];
            for (int k = 0; k < items.size(); k++) sz[k] = val(d, items.get(k));
            List<Rectangle2D> sub = slice(sz, rs.get(g));
            for (int k = 0; k < items.size(); k++) {
                Rectangle2D r = sub.get(k);
                cv.fillRect(r.getX(), r.getY(), r.getWidth(), r.getHeight(), accent(g));
                cv.stroke(r, new PdfCanvas.Stroke(0.75f), Color.WHITE);
                int i = items.get(k);
                String lab = i < leaf.size() && leaf.get(i) != null ? leaf.get(i) : "";
                text(cv, lab, r.getX() + 3, r.getY() + 3, SIZE, Color.WHITE, -1, -1);
            }
        }
    }

    /** 最上位の階層ごとの点の番号 (空のセルは上の行と同じ親)。 */
    private static Map<String, List<Integer>> groups(Data d) {
        Map<String, List<Integer>> out = new LinkedHashMap<>();
        List<String> top = d.levels.isEmpty() ? List.of() : d.levels.get(0);
        String last = "";
        for (int i = 0; i < d.vals.size(); i++) {
            String t = i < top.size() && top.get(i) != null ? top.get(i) : last;
            last = t;
            out.computeIfAbsent(t, k -> new ArrayList<>()).add(i);
        }
        return out;
    }

    private static double val(Data d, int i) { Double v = d.vals.get(i); return v == null || v < 0 ? 0 : v; }

    /** 値に比例して矩形を分ける (長い辺の向きに 1 列に並べる)。 */
    private static List<Rectangle2D> slice(double[] sizes, Rectangle2D r) {
        List<Rectangle2D> out = new ArrayList<>();
        double tot = 0;
        for (double s : sizes) tot += s;
        boolean horiz = r.getWidth() >= r.getHeight();
        double pos = 0;
        for (double s : sizes) {
            double f = tot > 0 ? s / tot : 0;
            if (horiz) out.add(new Rectangle2D.Double(r.getX() + pos * r.getWidth(), r.getY(), f * r.getWidth(), r.getHeight()));
            else out.add(new Rectangle2D.Double(r.getX(), r.getY() + pos * r.getHeight(), r.getWidth(), f * r.getHeight()));
            pos += f;
        }
        return out;
    }

    /** サンバースト: 内側の輪が最上位の階層、外側が葉。角度は値に比例し、12 時から時計回り。 */
    private void sunburst(PdfCanvas cv, Data d, Rectangle2D box) throws IOException {
        Map<String, List<Integer>> groups = groups(d);
        double tot = 0;
        for (int i = 0; i < d.vals.size(); i++) tot += val(d, i);
        if (tot <= 0) return;
        double rad = Math.min(box.getWidth(), box.getHeight()) / 2, cx = box.getCenterX(), cy = box.getCenterY();
        // 中央に穴を空け、穴と各階層の輪を同じ幅にする (2 階層で半径の 1/3 ずつ)
        int lv = Math.max(1, d.levels.size());
        double r0 = rad / (lv + 1), r1 = lv > 1 ? rad * 2 / (lv + 1) : rad, start = 90;
        List<String> leaf = d.levels.isEmpty() ? List.of() : d.levels.get(d.levels.size() - 1);
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.75f);
        int g = 0;
        for (Map.Entry<String, List<Integer>> e : groups.entrySet()) {
            double gs = 0;
            for (int i : e.getValue()) gs += val(d, i);
            double ext = -gs / tot * 360;
            Path2D inner = new Path2D.Double();
            inner.append(new Arc2D.Double(cx - r1, cy - r1, 2 * r1, 2 * r1, start, ext, Arc2D.OPEN), false);
            inner.append(new Arc2D.Double(cx - r0, cy - r0, 2 * r0, 2 * r0, start + ext, -ext, Arc2D.OPEN), true);
            inner.closePath();
            cv.fill(inner, accent(g));
            cv.stroke(inner, st, Color.WHITE);
            double mid = Math.toRadians(start + ext / 2);
            text(cv, e.getKey(), cx + Math.cos(mid) * (r0 + r1) / 2, cy - Math.sin(mid) * (r0 + r1) / 2, SIZE, Color.WHITE, 0, 0);
            double a = start;
            if (d.levels.size() > 1) {
                for (int i : e.getValue()) {
                    double x = -val(d, i) / tot * 360;
                    Path2D p = new Path2D.Double();
                    p.append(new Arc2D.Double(cx - rad, cy - rad, 2 * rad, 2 * rad, a, x, Arc2D.OPEN), false);
                    p.append(new Arc2D.Double(cx - r1, cy - r1, 2 * r1, 2 * r1, a + x, -x, Arc2D.OPEN), true);
                    p.closePath();
                    cv.fill(p, accent(g));
                    cv.stroke(p, st, Color.WHITE);
                    double m = Math.toRadians(a + x / 2), rr = (rad + r1) / 2;
                    String lab = i < leaf.size() && leaf.get(i) != null ? leaf.get(i) : "";
                    text(cv, lab, cx + Math.cos(m) * rr, cy - Math.sin(m) * rr, SIZE, Color.WHITE, 0, 0);
                    a += x;
                }
            }
            start += ext;
            g++;
        }
    }

    /** ヒストグラムの区間の数の上限。 */
    static final int MAX_BINS = 1000;
}
