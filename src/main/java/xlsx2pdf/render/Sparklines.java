package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.Path2D;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtension;
import org.w3c.dom.Element;

/**
 * スパークライン (Excel 2010 以降の拡張 x14:sparklineGroups)。折れ線・縦棒・勝敗を、セルの中に描く。
 * 寸法は Excel の印刷の実測: 描く範囲はセルの左右 1.5pt・上下 3pt の内側、点は範囲を等分した区画の中央、縦棒の幅は区画の 72%。
 */
public final class Sparklines {
    /** スパークラインの拡張の URI。 */
    static final String URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    static final double INSET_X = 1.5, INSET_Y = 3, BAR = 0.72, MARKER = 1.4;

    /** グループの書式。 */
    private static final class Group {
        String type = "line", empty = "zero", minType = "individual", maxType = "individual";
        double lineWeight = 0.75;
        Double manualMin, manualMax;
        boolean markers, high, low, first, last, negative, axis, rtl;
        Color series = new Color(0x376092), neg = new Color(0xD00000), axisColor = Color.BLACK, mark = new Color(0xD00000),
                cFirst = new Color(0xD00000), cLast = new Color(0xD00000), cHigh = new Color(0xD00000), cLow = new Color(0xD00000);
        List<Spark> items = new ArrayList<>();
        /** グループで揃える最小・最大 (計算後)。 */
        double gmin = Double.NaN, gmax = Double.NaN;
    }

    /** 1 つのスパークライン: 描くセルと値。 */
    private record Spark(int row, int col, List<Double> vals) {}

    private final RenderContext ctx;
    private final Map<Long, Spark> byCell = new HashMap<>();
    private final Map<Spark, Group> groupOf = new HashMap<>();

    private Sparklines(RenderContext ctx) { this.ctx = ctx; }

    /** シートのスパークラインを読む。なければ null。 */
    static Sparklines parse(RenderContext ctx, XSSFSheet sheet) {
        var ws = sheet.getCTWorksheet();
        boolean has = false;
        if (ws.isSetExtLst()) for (CTExtension e : ws.getExtLst().getExtArray()) has |= URI.equalsIgnoreCase(e.getUri());
        if (!has) return null;
        Element root;
        try (java.io.InputStream in = sheet.getPackagePart().getInputStream()) {
            root = Xml.parse(in);
        } catch (Exception e) {
            Warnings.add("スパークラインを読めませんでした", e);
            return null;
        }
        Sparklines out = new Sparklines(ctx);
        for (Element ext : Xml.children(Xml.child(root, "extLst"), "ext")) {
            if (!URI.equalsIgnoreCase(ext.getAttribute("uri"))) continue;
            for (Element gs : Xml.children(ext, "sparklineGroups")) {
                for (Element ge : Xml.children(gs, "sparklineGroup")) {
                    Group g = out.group(ge);
                    for (Element sl : Xml.children(Xml.child(ge, "sparklines"), "sparkline")) {
                        try {
                            String f = Xml.text(Xml.child(sl, "f")), at = Xml.text(Xml.child(sl, "sqref")).trim();
                            CellReference cr = new CellReference(at);
                            Spark sp = new Spark(cr.getRow(), cr.getCol(), out.values(sheet, f));
                            g.items.add(sp);
                            out.byCell.put((long) sp.row() * 16384 + sp.col(), sp);
                            out.groupOf.put(sp, g);
                        } catch (RuntimeException e) {
                            Warnings.add("スパークラインを読めませんでした", e);
                        }
                    }
                }
            }
        }
        return out.byCell.isEmpty() ? null : out;
    }

    /** グループの属性と色を読む。 */
    private Group group(Element e) {
        Group g = new Group();
        if (e.hasAttribute("type")) g.type = e.getAttribute("type");
        if (e.hasAttribute("displayEmptyCellsAs")) g.empty = e.getAttribute("displayEmptyCellsAs");
        g.lineWeight = Xml.doubleAttr(e, "lineWeight", 0.75);
        g.markers = Xml.boolAttr(e, "markers", false);
        g.high = Xml.boolAttr(e, "high", false);
        g.low = Xml.boolAttr(e, "low", false);
        g.first = Xml.boolAttr(e, "first", false);
        g.last = Xml.boolAttr(e, "last", false);
        g.negative = Xml.boolAttr(e, "negative", false);
        g.axis = Xml.boolAttr(e, "displayXAxis", false);
        g.rtl = Xml.boolAttr(e, "rightToLeft", false);
        if (e.hasAttribute("minAxisType")) g.minType = e.getAttribute("minAxisType");
        if (e.hasAttribute("maxAxisType")) g.maxType = e.getAttribute("maxAxisType");
        if (e.hasAttribute("manualMin")) g.manualMin = Xml.doubleAttr(e, "manualMin", 0);
        if (e.hasAttribute("manualMax")) g.manualMax = Xml.doubleAttr(e, "manualMax", 0);
        Color c;
        if ((c = color(Xml.child(e, "colorSeries"))) != null) g.series = c;
        if ((c = color(Xml.child(e, "colorNegative"))) != null) g.neg = c;
        if ((c = color(Xml.child(e, "colorAxis"))) != null) g.axisColor = c;
        if ((c = color(Xml.child(e, "colorMarkers"))) != null) g.mark = c;
        if ((c = color(Xml.child(e, "colorFirst"))) != null) g.cFirst = c;
        if ((c = color(Xml.child(e, "colorLast"))) != null) g.cLast = c;
        if ((c = color(Xml.child(e, "colorHigh"))) != null) g.cHigh = c;
        if ((c = color(Xml.child(e, "colorLow"))) != null) g.cLow = c;
        return g;
    }

    /** x14 の色の要素 (rgb / theme / tint / indexed)。 */
    private Color color(Element e) {
        if (e == null) return null;
        try {
            var ct = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
            if (e.hasAttribute("rgb")) ct.setRgb(java.util.HexFormat.of().parseHex(e.getAttribute("rgb")));
            if (e.hasAttribute("theme")) ct.setTheme(Long.parseLong(e.getAttribute("theme")));
            if (e.hasAttribute("indexed")) ct.setIndexed(Long.parseLong(e.getAttribute("indexed")));
            if (e.hasAttribute("tint")) ct.setTint(Double.parseDouble(e.getAttribute("tint")));
            XSSFColor xc = XSSFColor.from(ct, ctx.wb.getStylesSource().getIndexedColors());
            if (xc.isThemed() && ctx.wb.getStylesSource().getTheme() != null) ctx.wb.getStylesSource().getTheme().inheritFromThemeAsRequired(xc);
            return ctx.theme.color(xc);
        } catch (RuntimeException ex) {
            return null;
        }
    }

    /** 範囲 (Sheet1!A1:G1) の数値 (数値でないセルは null)。 */
    private List<Double> values(XSSFSheet own, String ref) {
        List<Double> out = new ArrayList<>();
        AreaReference ar = new AreaReference(ref, SpreadsheetVersion.EXCEL2007);
        CellReference a = ar.getFirstCell(), b = ar.getLastCell();
        Sheet sh = a.getSheetName() == null ? own : ctx.wb.getSheet(a.getSheetName());
        if (sh == null) return out;
        int r0 = Math.min(a.getRow(), b.getRow()), r1 = Math.max(a.getRow(), b.getRow());
        int c0 = Math.min(a.getCol(), b.getCol()), c1 = Math.max(a.getCol(), b.getCol());
        if ((long) (r1 - r0 + 1) * (c1 - c0 + 1) > 100_000) return out;
        for (int r = r0; r <= r1; r++) {
            Row row = sh.getRow(r);
            for (int c = c0; c <= c1; c++) {
                Cell cell = row == null ? null : row.getCell(c);
                Double v = null;
                if (cell != null) {
                    CellType t = cell.getCellType();
                    if (t == CellType.FORMULA) t = cell.getCachedFormulaResultType();
                    if (t == CellType.NUMERIC) v = cell.getNumericCellValue();
                }
                out.add(v);
            }
        }
        return out;
    }

    /** セル (r, c) にスパークラインがあれば、矩形 (x, y, w, h) に描く。 */
    void paint(PdfCanvas cv, int r, int c, double x, double y, double w, double h) throws IOException {
        Spark sp = byCell.get((long) r * 16384 + c);
        if (sp == null) return;
        Group g = groupOf.get(sp);
        List<Double> vals = new ArrayList<>(sp.vals());
        if (g.rtl) java.util.Collections.reverse(vals);
        int n = vals.size();
        if (n == 0) return;
        if (g.empty.equals("zero")) for (int i = 0; i < n; i++) if (vals.get(i) == null) vals.set(i, 0.0);
        double min = Double.MAX_VALUE, max = -Double.MAX_VALUE;
        int iMin = -1, iMax = -1;
        for (int i = 0; i < n; i++) {
            Double v = vals.get(i);
            if (v == null) continue;
            if (v < min) { min = v; iMin = i; }
            if (v > max) { max = v; iMax = i; }
        }
        if (iMin < 0) return;
        // 縦の範囲: 個別、グループで共通、指定値
        if (g.minType.equals("group") || g.maxType.equals("group")) groupRange(g);
        double lo = g.minType.equals("custom") && g.manualMin != null ? g.manualMin : g.minType.equals("group") ? g.gmin : min;
        double hi = g.maxType.equals("custom") && g.manualMax != null ? g.manualMax : g.maxType.equals("group") ? g.gmax : max;
        boolean column = g.type.equals("column");
        if (column) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
        if (hi <= lo) { hi = lo + 1; lo -= 0; }
        double ix = x + INSET_X, iw = w - 2 * INSET_X, iy = y + INSET_Y, ih = h - 2 * INSET_Y;
        if (iw <= 0 || ih <= 0) return;
        double slot = iw / n;
        final double flo = lo, fhi = hi;
        java.util.function.DoubleUnaryOperator py = v -> iy + ih - (v - flo) / (fhi - flo) * ih;
        cv.save();
        cv.clipRect(x, y, w, h);
        if (g.type.equals("stacked")) {
            // 勝敗: 正は上半分、負は下半分
            double mid = iy + ih / 2;
            for (int i = 0; i < n; i++) {
                Double v = vals.get(i);
                if (v == null || v == 0) continue;
                double bx = ix + i * slot + slot * (1 - BAR) / 2;
                Color col = pointColor(g, i, n, v, iMin, iMax, g.negative && v < 0 ? g.neg : g.series);
                if (v > 0) cv.fillRect(bx, iy, slot * BAR, ih / 2, col);
                else cv.fillRect(bx, mid, slot * BAR, ih / 2, col);
            }
        } else if (column) {
            double zy = py.applyAsDouble(0);
            for (int i = 0; i < n; i++) {
                Double v = vals.get(i);
                if (v == null) continue;
                double bx = ix + i * slot + slot * (1 - BAR) / 2, vy = py.applyAsDouble(v);
                Color col = pointColor(g, i, n, v, iMin, iMax, g.negative && v < 0 ? g.neg : g.series);
                cv.fillRect(bx, Math.min(zy, vy), slot * BAR, Math.abs(vy - zy), col);
            }
            if (g.axis) cv.fillRect(ix, zy - 0.375, iw, 0.75, g.axisColor);
        } else {
            if (g.axis && lo < 0 && hi > 0) cv.fillRect(ix, py.applyAsDouble(0) - 0.375, iw, 0.75, g.axisColor);
            Path2D p = new Path2D.Double();
            boolean started = false;
            for (int i = 0; i < n; i++) {
                Double v = vals.get(i);
                if (v == null) { if (!g.empty.equals("span")) started = false; continue; }
                double px = ix + (i + 0.5) * slot, yy = py.applyAsDouble(v);
                if (!started) { p.moveTo(px, yy); started = true; } else p.lineTo(px, yy);
            }
            PdfCanvas.Stroke st = new PdfCanvas.Stroke((float) g.lineWeight);
            st.cap = 1;
            st.join = 1;
            cv.stroke(p, st, g.series);
            for (int i = 0; i < n; i++) {
                Double v = vals.get(i);
                if (v == null) continue;
                Color mc = null;
                if (g.markers) mc = g.mark;
                if (g.negative && v < 0) mc = g.neg;
                if (g.first && i == 0) mc = g.cFirst;
                if (g.last && i == n - 1) mc = g.cLast;
                if (g.low && i == iMin) mc = g.cLow;
                if (g.high && i == iMax) mc = g.cHigh;
                if (mc == null) continue;
                double px = ix + (i + 0.5) * slot, yy = py.applyAsDouble(v), m = MARKER;
                Path2D d = new Path2D.Double();
                d.moveTo(px, yy - m); d.lineTo(px + m, yy); d.lineTo(px, yy + m); d.lineTo(px - m, yy); d.closePath();
                cv.fill(d, mc);
            }
        }
        cv.restore();
    }

    /** 縦棒・勝敗の点の色 (強調する点の指定が優先)。 */
    private static Color pointColor(Group g, int i, int n, double v, int iMin, int iMax, Color dflt) {
        if (g.high && i == iMax) return g.cHigh;
        if (g.low && i == iMin) return g.cLow;
        if (g.first && i == 0) return g.cFirst;
        if (g.last && i == n - 1) return g.cLast;
        return dflt;
    }

    /** グループで共通の最小・最大。 */
    private static void groupRange(Group g) {
        if (!Double.isNaN(g.gmin)) return;
        double mn = Double.MAX_VALUE, mx = -Double.MAX_VALUE;
        for (Spark s : g.items) for (Double v : s.vals()) if (v != null) { mn = Math.min(mn, v); mx = Math.max(mx, v); }
        g.gmin = mn == Double.MAX_VALUE ? 0 : mn;
        g.gmax = mx == -Double.MAX_VALUE ? 1 : mx;
    }

    /** シートのスパークラインのセルの一覧 (ページの準備用)。 */
    public record Cells(List<int[]> list) {
        /** ブロックのどれかにスパークラインのセルがあるか。 */
        public boolean any(List<Paginator.Segment> segs) {
            for (int[] rc : list) for (Paginator.Segment s : segs)
                if (s.body && rc[0] >= s.r0 && rc[0] <= s.r1 && rc[1] >= s.c0 && rc[1] <= s.c1) return true;
            return false;
        }
    }

    /** シートのスパークラインのセル (なければ空)。 */
    public static Cells cellsOf(RenderContext ctx, XSSFSheet sheet) {
        try {
            Sparklines sp = parse(ctx, sheet);
            return new Cells(sp == null ? List.of() : sp.cells());
        } catch (RuntimeException e) {
            return new Cells(List.of());
        }
    }

    /** スパークラインのあるセルか (白紙の判定など)。 */
    boolean has(int r, int c) { return byCell.containsKey((long) r * 16384 + c); }

    /** セルの一覧 ({行, 列})。 */
    List<int[]> cells() {
        List<int[]> out = new ArrayList<>();
        for (Spark s : byCell.values()) out.add(new int[] {s.row(), s.col()});
        return out;
    }
}
