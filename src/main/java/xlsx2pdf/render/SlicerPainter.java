package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.w3c.dom.Element;

/**
 * スライサー (sle:slicer) を描く: 枠、見出しと下線、項目のボタン。項目はテーブルの列 (tableSlicerCache) の値、または
 * ピボットキャッシュの共有項目 (tabular の i)。選んだ項目はアクセントの色の明るさ 60% 上げ、データのない項目は 80% 上げで塗る
 * (Excel の印刷の実測。スタイル SlicerStyleLightN / DarkN / OtherN の N 番目のアクセント)。
 */
final class SlicerPainter {
    private static final String REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer";
    private static final String REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";

    /** 見出しの下線までの高さ、最初のボタンの上端、左右の余白、ボタンの間隔、文字の左の余白 (pt。Excel の印刷の実測)。 */
    static final double CAPTION_LINE = 20.6, FIRST_BUTTON = 24.9, INSET = 7.2, BUTTON_GAP = 2.0, TEXT_PAD = 6.5;
    /** 文字の大きさ (pt)。 */
    static final float FONT = 11f;

    private final RenderContext ctx;
    private final XSSFSheet sheet;

    SlicerPainter(RenderContext ctx, XSSFSheet sheet) {
        this.ctx = ctx;
        this.sheet = sheet;
    }

    /** 1 つの項目: 表示の文字、選ばれているか、データがあるか。 */
    record Item(String text, boolean selected, boolean hasData) {}

    /** 名前 name のスライサーを矩形 r に描く。 */
    void paint(PdfCanvas cv, String name, Rectangle2D r) throws Exception {
        Element sl = findSlicer(name);
        if (sl == null) return;
        String caption = sl.hasAttribute("caption") ? sl.getAttribute("caption") : name;
        boolean showCaption = Xml.boolAttr(sl, "showCaption", true);
        double rowH = Xml.longAttr(sl, "rowHeight", 241300) / 12700.0;
        int cols = Math.max(1, Xml.intAttr(sl, "columnCount", 1));
        String style = Xml.attr(sl, "style", "SlicerStyleLight1");
        Element cache = findCache(sl.getAttribute("cache"));
        List<Item> items = cache == null ? List.of() : items(cache);

        // 色: スタイルの番号のアクセント
        int n = 1;
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("(\\d+)$").matcher(style);
        if (m.find()) n = Math.max(1, Math.min(6, Integer.parseInt(m.group(1))));
        Color accent = ctx.theme.scheme.getOrDefault("accent" + n, new Color(0x156082));
        Color sel = tint(accent, 0.6), noData = tint(accent, 0.8);

        cv.save();
        cv.clipRect(r.getX(), r.getY(), r.getWidth(), r.getHeight());
        cv.fillRect(r.getX(), r.getY(), r.getWidth(), r.getHeight(), Color.WHITE);
        FontManager.Spec font = ctx.fonts.spec(ctx.theme.minorLatin, ctx.theme.fontRef("+mn-ea", true), false, false);
        FontManager.Spec bold = ctx.fonts.spec(ctx.theme.minorLatin, ctx.theme.fontRef("+mn-ea", true), true, false);
        double x0 = r.getX() + INSET, w = r.getWidth() - 2 * INSET;
        double y = r.getY();
        if (showCaption) {
            double base = r.getY() + CAPTION_LINE - 5.5;
            cv.text(caption, x0, base, bold, FONT, Color.BLACK, 1, 0);
            cv.fillRect(x0, r.getY() + CAPTION_LINE - 0.35, w, 0.7, accent);
            y += FIRST_BUTTON;
        } else y += INSET;
        double bw = (w - (cols - 1) * BUTTON_GAP) / cols;
        for (int i = 0; i < items.size(); i++) {
            Item it = items.get(i);
            double bx = x0 + (i % cols) * (bw + BUTTON_GAP), by = y + (i / cols) * (rowH + BUTTON_GAP);
            if (by > r.getMaxY()) break;
            Color fill = it.selected ? (it.hasData ? sel : noData) : Color.WHITE;
            Color border = it.selected && it.hasData ? new Color(0x999999) : new Color(0xCCCCCC);
            cv.fillRect(bx, by, bw, rowH, fill);
            cv.stroke(new Rectangle2D.Double(bx + 0.35, by + 0.35, bw - 0.7, rowH - 0.7), new PdfCanvas.Stroke(0.7f), it.hasData ? border : tint(border, 0.5));
            Color tc = it.hasData ? Color.BLACK : new Color(0xA6A6A6);
            double base = by + rowH / 2 + (font.ascent - font.descent) * FONT / 2;
            cv.save();
            cv.clipRect(bx, by, bw, rowH);
            cv.text(it.text, bx + TEXT_PAD, base, font, FONT, tc, 1, 0);
            cv.restore();
        }
        cv.stroke(new Rectangle2D.Double(r.getX() + 0.35, r.getY() + 0.35, r.getWidth() - 0.7, r.getHeight() - 0.7), new PdfCanvas.Stroke(0.7f), accent);
        cv.restore();
    }

    /** 色の明るさ (HSL の L) を f の割合だけ白に近づける (lumMod 1−f、lumOff f)。 */
    static Color tint(Color c, double f) {
        float[] hsl = hsl(c);
        double l = hsl[2] * (1 - f) + f;
        return fromHsl(hsl[0], hsl[1], l);
    }

    private static float[] hsl(Color c) {
        double r = c.getRed() / 255.0, g = c.getGreen() / 255.0, b = c.getBlue() / 255.0;
        double max = Math.max(r, Math.max(g, b)), min = Math.min(r, Math.min(g, b)), l = (max + min) / 2, h = 0, s = 0;
        if (max != min) {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
            else if (max == g) h = (b - r) / d + 2;
            else h = (r - g) / d + 4;
            h /= 6;
        }
        return new float[] {(float) h, (float) s, (float) l};
    }

    private static Color fromHsl(double h, double s, double l) {
        double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
        return new Color(ch(hue(p, q, h + 1.0 / 3)), ch(hue(p, q, h)), ch(hue(p, q, h - 1.0 / 3)));
    }

    private static int ch(double v) { return (int) Math.round(Math.max(0, Math.min(1, v)) * 255); }

    private static double hue(double p, double q, double t) {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 0.5) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    // ------------------------------------------------------------ 定義の読み込み

    private static Element parse(PackagePart p) throws Exception {
        try (InputStream in = p.getInputStream()) {
            return Xml.parse(in);
        }
    }

    /** シートのスライサーの定義 (slicers/slicer) を名前で探す。 */
    private Element findSlicer(String name) throws Exception {
        PackagePart sp = sheet.getPackagePart();
        for (PackageRelationship rel : sp.getRelationshipsByType(REL_SLICER)) {
            Element root = parse(sp.getRelatedPart(rel));
            for (Element s : Xml.children(root, "slicer")) if (name.equals(s.getAttribute("name"))) return s;
        }
        return null;
    }

    /** ブックのスライサーキャッシュを名前で探す。 */
    private Element findCache(String name) throws Exception {
        PackagePart wp = sheet.getWorkbook().getPackagePart();
        for (PackageRelationship rel : wp.getRelationshipsByType(REL_SLICER_CACHE)) {
            Element root = parse(wp.getRelatedPart(rel));
            if (name.equals(root.getAttribute("name"))) return root;
        }
        return null;
    }

    /** キャッシュの項目。データのある項目を先に並べる (Excel の既定 crossFilter="showItemsWithDataAtTop")。 */
    private List<Item> items(Element cache) throws Exception {
        List<Item> out = new ArrayList<>();
        Element tab = Xml.path(cache, "data", "tabular");
        if (tab != null) out = pivotItems(cache, tab);
        else {
            for (Element ext : Xml.children(Xml.child(cache, "extLst"), "ext")) {
                Element t = Xml.child(ext, "tableSlicerCache");
                if (t != null) out = tableItems(Xml.intAttr(t, "tableId", -1), Xml.intAttr(t, "column", 1),
                        !"descending".equals(t.getAttribute("sortOrder")), t.getAttribute("crossFilter"));
            }
        }
        return out;
    }

    /** テーブルの列の値。選択は列の絞り込み (autoFilter の filterColumn)、データの有無は非表示でない行にあるか。 */
    private List<Item> tableItems(int tableId, int column, boolean ascending, String crossFilter) {
        for (org.apache.poi.ss.usermodel.Sheet sh : sheet.getWorkbook()) {
            for (XSSFTable t : ((XSSFSheet) sh).getTables()) {
                if (t.getCTTable().getId() != tableId) continue;
                int c = t.getStartColIndex() + column - 1;
                int r0 = t.getStartRowIndex() + Math.max(1, (int) t.getCTTable().getHeaderRowCount());
                int r1 = t.getEndRowIndex() - (int) t.getCTTable().getTotalsRowCount();
                Set<String> filter = null;
                var af = t.getCTTable().getAutoFilter();
                if (af != null) {
                    for (var fc : af.getFilterColumnArray()) {
                        if (fc.getColId() == column - 1 && fc.isSetFilters()) {
                            filter = new LinkedHashSet<>();
                            for (var f : fc.getFilters().getFilterArray()) filter.add(f.getVal());
                        }
                    }
                }
                DataFormatter df = new DataFormatter();
                Map<String, Object> keys = new LinkedHashMap<>();
                Set<String> visible = new LinkedHashSet<>();
                for (int r = r0; r <= r1; r++) {
                    Row row = sh.getRow(r);
                    Cell cell = row == null ? null : row.getCell(c);
                    String s = cell == null ? "" : text(cell, df);
                    Object key = cell != null && (cell.getCellType() == CellType.NUMERIC
                            || cell.getCellType() == CellType.FORMULA && cell.getCachedFormulaResultType() == CellType.NUMERIC)
                            ? (Object) cell.getNumericCellValue() : s.toLowerCase(java.util.Locale.ROOT);
                    keys.putIfAbsent(s, key);
                    if (row == null || !row.getZeroHeight()) visible.add(s);
                }
                List<String> vals = new ArrayList<>(keys.keySet());
                vals.sort((a, b) -> {
                    Object ka = keys.get(a), kb = keys.get(b);
                    int cmp;
                    if (ka instanceof Double x && kb instanceof Double y) cmp = Double.compare(x, y);
                    else if (ka instanceof Double) cmp = -1;
                    else if (kb instanceof Double) cmp = 1;
                    else cmp = ((String) ka).compareTo((String) kb);
                    return ascending ? cmp : -cmp;
                });
                List<Item> withData = new ArrayList<>(), without = new ArrayList<>();
                for (String v : vals) {
                    if (v.isEmpty()) continue;
                    boolean sel = filter == null || filter.contains(v);
                    boolean data = !sel || visible.contains(v);
                    (data || "showItemsWithNoData".equals(crossFilter) ? withData : without).add(new Item(v, sel, data));
                }
                withData.addAll(without);
                return withData;
            }
        }
        return List.of();
    }

    private static String text(Cell cell, DataFormatter df) {
        if (cell.getCellType() == CellType.FORMULA) {
            return switch (cell.getCachedFormulaResultType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> df.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(),
                        cell.getCellStyle().getDataFormatString());
                case BOOLEAN -> cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                default -> "";
            };
        }
        return df.formatCellValue(cell);
    }

    /** ピボットのスライサー: tabular の i (x: 共有項目の番号、s: 選択、nd: データなし)。 */
    private List<Item> pivotItems(Element cache, Element tab) throws Exception {
        List<String> shared = sharedItems(Xml.intAttr(tab, "pivotCacheId", -1), cache.getAttribute("sourceName"));
        List<Item> withData = new ArrayList<>(), without = new ArrayList<>();
        Element items = Xml.child(tab, "items");
        for (Element i : Xml.children(items, "i")) {
            int x = Xml.intAttr(i, "x", -1);
            String t = x >= 0 && x < shared.size() ? shared.get(x) : "";
            boolean sel = Xml.boolAttr(i, "s", false), nd = Xml.boolAttr(i, "nd", false);
            (nd ? without : withData).add(new Item(t, sel, !nd));
        }
        withData.addAll(without);
        return withData;
    }

    /** ピボットキャッシュ (ブックの pivotCaches の cacheId) のフィールドの共有項目の文字列。 */
    private List<String> sharedItems(int cacheId, String field) throws Exception {
        List<String> out = new ArrayList<>();
        PackagePart wp = sheet.getWorkbook().getPackagePart();
        Element wb = parse(wp);
        String rid = null;
        for (Element pc : Xml.children(Xml.child(wb, "pivotCaches"), "pivotCache")) {
            if (Xml.intAttr(pc, "cacheId", -2) == cacheId) rid = pc.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        }
        // Excel 2010 以降のスライサー用のキャッシュは extLst の x14:pivotCaches にある
        if (rid == null) {
            for (Element ext : Xml.children(Xml.child(wb, "extLst"), "ext")) {
                for (Element pc : Xml.children(Xml.child(ext, "pivotCaches"), "pivotCache")) {
                    if (Xml.intAttr(pc, "cacheId", -2) == cacheId) rid = pc.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                }
            }
        }
        if (rid == null || rid.isEmpty()) return out;
        PackageRelationship rel = wp.getRelationship(rid);
        if (rel == null) return out;
        Element def = parse(wp.getRelatedPart(rel));
        for (Element cf : Xml.children(Xml.child(def, "cacheFields"), "cacheField")) {
            if (!field.equals(cf.getAttribute("name"))) continue;
            for (Element v : Xml.children(Xml.child(cf, "sharedItems"))) {
                String t = v.getLocalName().equals("m") ? "(空白)" : v.getAttribute("v");
                out.add(t);
            }
        }
        return out;
    }
}
