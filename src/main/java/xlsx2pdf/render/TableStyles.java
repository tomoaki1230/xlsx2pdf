package xlsx2pdf.render;

import java.awt.Color;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DifferentialStyleProvider;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.TableStyle;
import org.apache.poi.ss.usermodel.TableStyleInfo;
import org.apache.poi.ss.usermodel.TableStyleType;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;

import xlsx2pdf.render.CondFormat.Edge;

/**
 * シートのテーブル (挿入 → テーブル) とピボットテーブルのスタイルを、セルごとの塗り・文字・罫線にする。定義は組み込み
 * (POI の presetTableStyles.xml) とブックの独自スタイル。要素は優先順位の低い順に重ね、上の要素が下の要素を上書きする
 * (文字の太字・斜体などは、上の要素で明示した val="0" も上書きする)。セルの書式はテーブルより優先する (呼び出し側で扱う)。
 * 縞は見えている行・列で数える。
 * <p>
 * テーブルの順: テーブル全体、列の縞、行の縞、最初の列、最後の列、見出し行、集計行、見出しの最初・最後のセル、集計行の最初・最後のセル。
 * <p>
 * ピボットの順: テーブル全体、レポートフィルターの名前・値、列の縞、行の縞 (縞はデータの範囲)、
 * 最初の列 (行ラベルの列)、見出し行、左上の見出しのセル、小計の列 1〜3、空白行、小計の行 1〜3、列の小見出し 1〜3、
 * 行の小見出し 1〜3、総計の列、総計の行。
 */
final class TableStyles {
    /** セルに効くテーブルの塗りと文字 (なければ null)。 */
    static final class Look {
        Color fill, font;
        /** 要素が決めた太字・斜体・取り消し線 (null は決めていない)。 */
        Boolean bold, italic, strike;
        /** 要素が決めた下線 (null は決めていない。0 なし、1 一重、2 二重)。 */
        Integer underline;
        /** ピボットの行ラベルのセル: 字下げを展開ボタンの幅とピボットの段の幅で数える。 */
        boolean pivotLabel;
    }

    /** ピボットの行ラベルの、展開ボタンの分の字下げと 1 段の字下げ (画面の px)。Excel の印刷の実測 (12.45pt、13.85pt)。 */
    static final double PIVOT_BUTTON_PX = 16.6, PIVOT_LEVEL_PX = 18.47;

    /** テーブルの 1 要素と、その要素の範囲の中でのセルの位置 (外周の辺か)。 */
    private record Layer(DifferentialStyleProvider p, boolean top, boolean right, boolean bottom, boolean left, int rank) {}

    /** 縞の 1 行 (1 列)。kind は 0 なし、1 第 1 の縞、2 第 2 の縞。first・last は縞の端か。 */
    private record Band(int kind, boolean first, boolean last) {}

    /** ピボットの行 (列) の種類。 */
    private static final int ITEM = 0, SUBHEAD = 1, SUBTOTAL = 2, BLANK = 3, GRAND = 4;

    private static final class Table {
        int r0, c0, r1, c1, bodyR0, bodyR1;
        TableStyle style;
        TableStyleInfo info;
        Band[] rowBands, colBands;
        boolean pivot;
        // ---- ピボット ----
        /** 列の見出しの最初の行、データの最初の列。 */
        int hdrR0, dataC0;
        /** 本体の行 (bodyR0 から) の種類と段 (1 始まり)。 */
        int[] rowKind, rowLevel;
        /** 本体の行の項目の r (前の行と同じ項目の数)。 */
        int[] itemR;
        /** データの列 (dataC0 から) の種類と段。 */
        int[] colKind, colLevel;
        /** 行ラベルの列 (c0 から) の、表形式の項目の段 (0 はなし) と、展開ボタンの分の字下げをするか。 */
        int[] tabLevel;
        boolean[] button;
        int colFields;
        boolean showColHeaders = true;
        /** レポートフィルターの行の範囲 (なければ pageR0 > pageR1)。 */
        int pageR0 = 1, pageR1 = 0;
    }

    private final List<Table> tables = new ArrayList<>();
    private final Map<Long, Look> looks = new HashMap<>();
    private final Theme theme;

    TableStyles(XSSFSheet sheet, SheetLayout L, Theme theme) {
        this.theme = theme;
        for (XSSFTable x : sheet.getTables()) {
            try {
                addTable(x, L);
            } catch (RuntimeException e) {
                // 1 つの表が読めなくても、ほかの表のスタイルは読む
                Warnings.add("表のスタイルを読めませんでした", e);
            }
        }
        readPivots(sheet, L);
    }

    private void addTable(XSSFTable x, SheetLayout L) {
        {
            if (!styled(x)) return;
            AreaReference a = x.getCellReferences();
            if (a == null) return;
            Table t = new Table();
            t.r0 = a.getFirstCell().getRow(); t.c0 = a.getFirstCell().getCol();
            t.r1 = a.getLastCell().getRow(); t.c1 = a.getLastCell().getCol();
            t.bodyR0 = t.r0 + x.getHeaderRowCount();
            t.bodyR1 = t.r1 - x.getTotalsRowCount();
            t.info = x.getStyle();
            t.style = t.info.getStyle();
            t.rowBands = bands(t.style, TableStyleType.firstRowStripe, TableStyleType.secondRowStripe, t.bodyR0, t.bodyR1, true, L);
            t.colBands = bands(t.style, TableStyleType.firstColumnStripe, TableStyleType.secondColumnStripe, t.c0, t.c1, false, L);
            tables.add(t);
        }
    }

    private void readPivots(XSSFSheet sheet, SheetLayout L) {
        List<org.apache.poi.xssf.usermodel.XSSFPivotTable> pts;
        try {
            pts = sheet.getPivotTables();
        } catch (RuntimeException e) {
            Warnings.add("ピボットテーブルのスタイルを読めませんでした", e);
            pts = List.of();
        }
        for (org.apache.poi.xssf.usermodel.XSSFPivotTable pt : pts) {
            // 1 つのピボットが読めなくても、ほかのピボットのスタイルは読む
            try {
                Table t = pivot(pt, sheet, L);
                if (t != null) tables.add(t);
            } catch (RuntimeException e) {
                Warnings.add("ピボットテーブルのスタイルを読めませんでした", e);
            }
        }
    }

    /**
     * ピボットテーブル: 範囲 (location) の中で、firstDataRow までを見出し、firstDataCol までの列を行ラベルの列とする。
     * 行 (列) の種類は rowItems (colItems) の i の t (総計 grand、空白 blank、ほかは小計) と、段 = r + x の数 (行の項目の数より
     * 浅い項目の行は小見出しの行) で決める。
     */
    private Table pivot(org.apache.poi.xssf.usermodel.XSSFPivotTable pt, XSSFSheet sheet, SheetLayout L) {
        var def = pt.getCTPivotTableDefinition();
        if (!def.isSetPivotTableStyleInfo() || def.getPivotTableStyleInfo().getName() == null) return null;
        var si = def.getPivotTableStyleInfo();
        TableStyle st = sheet.getWorkbook().getStylesSource().getExplicitTableStyle(si.getName());
        if (st == null) st = builtinPivotStyle(si.getName());
        if (st == null) return null;
        AreaReference a = new AreaReference(def.getLocation().getRef(), org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
        Table t = new Table();
        t.pivot = true;
        t.r0 = a.getFirstCell().getRow(); t.c0 = a.getFirstCell().getCol();
        t.r1 = a.getLastCell().getRow(); t.c1 = a.getLastCell().getCol();
        var loc = def.getLocation();
        t.bodyR0 = Math.min(t.r1 + 1, t.r0 + (int) loc.getFirstDataRow());
        t.hdrR0 = t.r0 + (int) loc.getFirstHeaderRow();
        t.dataC0 = Math.min(t.c1 + 1, t.c0 + (int) loc.getFirstDataCol());
        // 行の項目
        int nRow = def.isSetRowFields() ? def.getRowFields().sizeOfFieldArray() : 0;
        t.rowKind = new int[t.r1 - t.bodyR0 + 1];
        t.rowLevel = new int[t.rowKind.length];
        t.itemR = new int[t.rowKind.length];
        boolean grandRow = false;
        if (def.isSetRowItems()) {
            var items = def.getRowItems().getIArray();
            for (int k = 0; k < items.length && k < t.rowKind.length; k++) {
                var it = items[k];
                int depth = (int) (it.isSetR() ? it.getR() : 0) + it.sizeOfXArray();
                String ty = it.isSetT() ? it.getT().toString() : "data";
                t.rowLevel[k] = Math.max(1, depth);
                t.itemR[k] = (int) (it.isSetR() ? it.getR() : 0);
                t.rowKind[k] = switch (ty) {
                    case "grand" -> GRAND;
                    case "blank" -> BLANK;
                    case "data" -> depth < nRow ? SUBHEAD : ITEM;
                    default -> SUBTOTAL;
                };
                if (t.rowKind[k] == GRAND) grandRow = true;
            }
        }
        int nCol = def.isSetColFields() ? def.getColFields().sizeOfFieldArray() : 0;
        t.colFields = nCol;
        t.colKind = new int[Math.max(0, t.c1 - t.dataC0 + 1)];
        t.colLevel = new int[t.colKind.length];
        if (def.isSetColItems()) {
            var items = def.getColItems().getIArray();
            for (int k = 0; k < items.length && k < t.colKind.length; k++) {
                var it = items[k];
                int depth = (int) (it.isSetR() ? it.getR() : 0) + it.sizeOfXArray();
                String ty = it.isSetT() ? it.getT().toString() : "data";
                t.colLevel[k] = Math.max(1, depth);
                t.colKind[k] = switch (ty) {
                    case "grand" -> GRAND;
                    case "blank" -> BLANK;
                    case "data" -> ITEM;
                    default -> SUBTOTAL;
                };
            }
        }
        // 総計の行は本体の後ろ (集計行の要素)
        t.bodyR1 = t.r1;
        if (grandRow) while (t.bodyR1 >= t.bodyR0 && t.rowKind[t.bodyR1 - t.bodyR0] == GRAND) t.bodyR1--;
        // 行ラベルの列: コンパクト形式の項目は次の項目と同じ列。表形式 (outline="0") の項目の列は行の小見出し。
        // 最後でない項目のある列は展開ボタンの分を字下げする (ボタンは印刷しないが、場所は空ける)
        int nLab = Math.max(0, t.dataC0 - t.c0);
        t.tabLevel = new int[nLab];
        t.button = new boolean[nLab];
        boolean defCompact = !def.isSetCompact() || def.getCompact();
        boolean drill = !def.isSetShowDrill() || def.getShowDrill();
        int col = 0;
        for (int k = 0; k < nRow; k++) {
            int fi = def.getRowFields().getFieldArray(k).getX();
            var pf = fi >= 0 && def.isSetPivotFields() && fi < def.getPivotFields().sizeOfPivotFieldArray()
                    ? def.getPivotFields().getPivotFieldArray(fi) : null;
            boolean outline = pf == null || !pf.isSetOutline() || pf.getOutline();
            boolean compact = defCompact && (pf == null || !pf.isSetCompact() || pf.getCompact()) && outline;
            if (col < nLab) {
                if (!outline && k < nRow - 1) t.tabLevel[col] = k + 1;
                if (drill && k < nRow - 1 && fi >= 0) t.button[col] = true;
            }
            if (!compact) col++;
        }
        t.style = st;
        final TableStyle fst = st;
        final boolean rs = si.getShowRowStripes(), cs = si.getShowColStripes(), rh = si.getShowRowHeaders(), lc = si.getShowLastColumn();
        t.showColHeaders = si.getShowColHeaders();
        final String nm = si.getName();
        t.info = new TableStyleInfo() {
            public boolean isShowColumnStripes() { return cs; }
            public boolean isShowRowStripes() { return rs; }
            public boolean isShowFirstColumn() { return rh; }
            public boolean isShowLastColumn() { return lc; }
            public String getName() { return nm; }
            public TableStyle getStyle() { return fst; }
        };
        // レポートフィルター: 範囲の 1 行上を空けて、その上に並ぶ (名前の列 c0、値の列 c0 + 1)
        int pages = def.isSetPageFields() ? def.getPageFields().sizeOfPageFieldArray() : 0;
        int perCol = loc.isSetRowPageCount() ? (int) loc.getRowPageCount() : pages;
        if (pages > 0 && perCol > 0) {
            t.pageR1 = t.r0 - 2;
            t.pageR0 = t.pageR1 - perCol + 1;
        }
        t.rowBands = bands(st, TableStyleType.firstRowStripe, TableStyleType.secondRowStripe, t.bodyR0, t.bodyR1, true, L);
        t.colBands = bands(st, TableStyleType.firstColumnStripe, TableStyleType.secondColumnStripe, t.dataC0, t.c1, false, L);
        return t;
    }

    private static final Map<String, TableStyle> PIVOT_STYLES = new java.util.concurrent.ConcurrentHashMap<>();

    /**
     * 組み込みのピボットのスタイル。POI の XSSFBuiltinTableStyle は定義の dxfId を 1 始まりとして読むが、ピボットのスタイルの
     * 定義 (presetTableStyles.xml) は 0 始まりなので 1 つずれる (総計の行に縞の灰色が当たる)。定義を読み直して 0 始まりで作る。
     */
    static TableStyle builtinPivotStyle(String name) {
        if (name == null || !name.startsWith("PivotStyle")) {
            try {
                return name == null ? null : org.apache.poi.xssf.usermodel.XSSFBuiltinTableStyle.valueOf(name).getStyle();
            } catch (IllegalArgumentException e) {
                return null;
            }
        }
        return PIVOT_STYLES.computeIfAbsent(name, n -> {
            try (java.io.InputStream in = org.apache.poi.xssf.usermodel.XSSFBuiltinTableStyle.class.getResourceAsStream("presetTableStyles.xml")) {
                if (in == null) return null;
                String xml = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
                int i = xml.indexOf("<" + n + ">"), j = xml.indexOf("</" + n + ">");
                if (i < 0 || j < 0) return null;
                String blk = xml.substring(i, j);
                // スタイルシートの文書に包んで読む (XMLBeans は要素だけの断片を読めない)
                int d0 = blk.indexOf("<dxfs"), d1 = blk.indexOf("</dxfs>") + 7;
                int t0 = blk.indexOf("<tableStyles"), t1 = blk.indexOf("</tableStyles>") + 14;
                if (t0 < 0 || t1 < 14) return null;
                String doc = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                        + (d0 >= 0 ? blk.substring(d0, d1) : "") + blk.substring(t0, t1) + "</styleSheet>";
                var ss = org.openxmlformats.schemas.spreadsheetml.x2006.main.StyleSheetDocument.Factory.parse(doc).getStyleSheet();
                var dxfs = ss.isSetDxfs() ? ss.getDxfs() : org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxfs.Factory.newInstance();
                var ct = ss.getTableStyles().getTableStyleArray(0);
                return new org.apache.poi.xssf.usermodel.XSSFTableStyle(0, dxfs, ct, new org.apache.poi.xssf.usermodel.DefaultIndexedColorMap());
            } catch (Exception e) {
                Warnings.add("組み込みのピボットのスタイル「" + n + "」を読めませんでした", e);
                return null;
            }
        });
    }

    /** スタイルが解決できるテーブルか (スタイルの指定がなければ何も描かない)。 */
    static boolean styled(XSSFTable x) {
        return x.getStyle() != null && x.getStyle().getStyle() != null;
    }

    /** 見えている行 (列) を順に、第 1 の縞と第 2 の縞に振り分ける。要素の size が縞の幅。 */
    private static Band[] bands(TableStyle st, TableStyleType first, TableStyleType second, int from, int to, boolean rows, SheetLayout L) {
        int n = Math.max(0, to - from + 1);
        Band[] out = new Band[n];
        int s1 = size(st.getStyle(first)), s2 = size(st.getStyle(second)), pos = 0;
        int period = s1 + s2;
        int[] vis = new int[n];
        int nv = 0;
        for (int i = 0; i < n; i++) if ((rows ? L.rh(from + i) : L.cw(from + i)) > 0) vis[nv++] = i;
        for (int v = 0; v < nv; v++) {
            int p = pos++ % period;
            boolean isFirst = p < s1;
            boolean head = p == 0 || p == s1;
            boolean tail = p == s1 - 1 || p == period - 1;
            out[vis[v]] = new Band(isFirst ? 1 : 2, head || v == 0, tail || v == nv - 1);
        }
        return out;
    }

    private static int size(DifferentialStyleProvider p) {
        return p == null || p.getStripeSize() <= 0 ? 1 : p.getStripeSize();
    }

    boolean any() { return !tables.isEmpty(); }

    private Table at(int r, int c) {
        for (Table t : tables) {
            if (r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1) return t;
            if (t.pivot && r >= t.pageR0 && r <= t.pageR1 && (c == t.c0 || c == t.c0 + 1)) return t;
        }
        return null;
    }

    /** 優先順位の低い順の要素。 */
    private List<Layer> layers(Table t, int r, int c) {
        return t.pivot ? pivotLayers(t, r, c) : tableLayers(t, r, c);
    }

    private static List<Layer> tableLayers(Table t, int r, int c) {
        List<Layer> out = new ArrayList<>(8);
        TableStyle s = t.style;
        boolean tl = c == t.c0, tr = c == t.c1;
        add(out, s, TableStyleType.wholeTable, r == t.r0, tr, r == t.r1, tl);
        boolean header = r < t.bodyR0, total = r > t.bodyR1;
        if (!header && !total) {
            if (t.info.isShowColumnStripes()) {
                Band b = t.colBands[c - t.c0];
                if (b != null) add(out, s, b.kind == 1 ? TableStyleType.firstColumnStripe : TableStyleType.secondColumnStripe,
                        r == t.bodyR0, b.last, r == t.bodyR1, b.first);
            }
            if (t.info.isShowRowStripes()) {
                Band b = t.rowBands[r - t.bodyR0];
                if (b != null) add(out, s, b.kind == 1 ? TableStyleType.firstRowStripe : TableStyleType.secondRowStripe,
                        b.first, tr, b.last, tl);
            }
        }
        if (t.info.isShowFirstColumn() && tl) add(out, s, TableStyleType.firstColumn, r == t.r0, true, r == t.r1, true);
        if (t.info.isShowLastColumn() && tr) add(out, s, TableStyleType.lastColumn, r == t.r0, true, r == t.r1, true);
        if (header) add(out, s, TableStyleType.headerRow, r == t.r0, tr, r == t.bodyR0 - 1, tl);
        if (total) add(out, s, TableStyleType.totalRow, r == t.bodyR1 + 1, tr, r == t.r1, tl);
        if (header && tl) add(out, s, TableStyleType.firstHeaderCell, true, true, true, true);
        if (header && tr) add(out, s, TableStyleType.lastHeaderCell, true, true, true, true);
        if (total && tl) add(out, s, TableStyleType.firstTotalCell, true, true, true, true);
        if (total && tr) add(out, s, TableStyleType.lastTotalCell, true, true, true, true);
        return out;
    }

    private static final TableStyleType[] SUBTOTAL_COL = {TableStyleType.firstSubtotalColumn, TableStyleType.secondSubtotalColumn, TableStyleType.thirdSubtotalColumn};
    private static final TableStyleType[] SUBTOTAL_ROW = {TableStyleType.firstSubtotalRow, TableStyleType.secondSubtotalRow, TableStyleType.thirdSubtotalRow};
    private static final TableStyleType[] COL_SUBHEAD = {TableStyleType.firstColumnSubheading, TableStyleType.secondColumnSubheading, TableStyleType.thirdColumnSubheading};
    private static final TableStyleType[] ROW_SUBHEAD = {TableStyleType.firstRowSubheading, TableStyleType.secondRowSubheading, TableStyleType.thirdRowSubheading};

    /** 段 (1 始まり) の要素。4 段より深いものは 3 段目と 2 段目を交互に使う (Excel と同じ)。 */
    private static TableStyleType level(TableStyleType[] a, int lv) {
        return lv <= 3 ? a[Math.max(0, lv - 1)] : a[(lv % 2 == 0) ? 1 : 2];
    }

    private static List<Layer> pivotLayers(Table t, int r, int c) {
        List<Layer> out = new ArrayList<>(10);
        TableStyle s = t.style;
        if (r < t.r0) {   // レポートフィルター
            boolean top = r == t.pageR0, bottom = r == t.pageR1;
            add(out, s, c == t.c0 ? TableStyleType.pageFieldLabels : TableStyleType.pageFieldValues, top, true, bottom, true);
            return out;
        }
        boolean tl = c == t.c0, tr = c == t.c1;
        boolean header = r < t.bodyR0, label = c < t.dataC0;
        int ri = r - t.bodyR0, ci = c - t.dataC0;
        int rk = header ? -1 : t.rowKind[ri], ck = label ? -1 : t.colKind[ci];
        add(out, s, TableStyleType.wholeTable, r == t.r0, tr, r == t.r1, tl);
        if (!header && r <= t.bodyR1 && !label) {
            if (t.info.isShowColumnStripes()) {
                Band b = t.colBands[ci];
                if (b != null) add(out, s, b.kind == 1 ? TableStyleType.firstColumnStripe : TableStyleType.secondColumnStripe,
                        r == t.bodyR0, b.last, r == t.bodyR1, b.first);
            }
        }
        if (!header && r <= t.bodyR1 && t.info.isShowRowStripes()) {
            Band b = t.rowBands[ri];
            if (b != null) add(out, s, b.kind == 1 ? TableStyleType.firstRowStripe : TableStyleType.secondRowStripe,
                    b.first, tr, b.last, tl);
        }
        if (t.info.isShowFirstColumn() && label && !header) add(out, s, TableStyleType.firstColumn, r == t.bodyR0, c == t.dataC0 - 1, r == t.r1, tl);
        if (header && t.showColHeaders) add(out, s, TableStyleType.headerRow, r == t.r0, tr, r == t.bodyR0 - 1, tl);
        if (r == t.r0 && label) add(out, s, TableStyleType.firstHeaderCell, true, c == t.dataC0 - 1, true, tl);
        if (ck == SUBTOTAL && r >= t.hdrR0) add(out, s, level(SUBTOTAL_COL, t.colLevel[ci]), r == t.hdrR0, true, r == t.r1, true);
        if (rk == BLANK) add(out, s, TableStyleType.blankRow, true, tr, true, tl);
        if (rk == SUBTOTAL) add(out, s, level(SUBTOTAL_ROW, t.rowLevel[ri]), true, tr, true, tl);
        if (header && !label && r >= t.hdrR0 && r - t.hdrR0 < t.colFields)
            add(out, s, level(COL_SUBHEAD, r - t.hdrR0 + 1), true, tr, true, c == t.dataC0);
        if (rk == SUBHEAD) add(out, s, level(ROW_SUBHEAD, t.rowLevel[ri]), true, tr, true, tl);
        if (label && rk == ITEM && t.tabLevel[c - t.c0] > 0) {
            // 表形式の項目の列: 項目が変わる行から、次に変わる行 (または小計の行) の前まで
            int lv = t.tabLevel[c - t.c0];
            add(out, s, level(ROW_SUBHEAD, lv), groupStart(t, ri, lv), true, groupEnd(t, ri, lv), true);
        }
        if (t.info.isShowLastColumn() && ck == GRAND && r >= t.hdrR0) add(out, s, TableStyleType.lastColumn, r == t.hdrR0, true, r == t.r1, true);
        if (!header && rk == GRAND) add(out, s, TableStyleType.totalRow, r == t.bodyR1 + 1, tr, r == t.r1, tl);
        return out;
    }

    /** 表形式の段 lv の項目の始まりの行か (行の項目の r が lv より小さければ、その段の項目が変わる)。 */
    private static boolean groupStart(Table t, int ri, int lv) {
        return ri == 0 || t.rowKind[ri - 1] != ITEM || itemR(t, ri) < lv;
    }

    private static boolean groupEnd(Table t, int ri, int lv) {
        return ri + 1 >= t.rowKind.length || t.rowKind[ri + 1] != ITEM || itemR(t, ri + 1) < lv;
    }

    /** 行の項目の r (前の行と同じ項目の数)。 */
    private static int itemR(Table t, int ri) {
        return t.itemR == null || ri >= t.itemR.length ? 0 : t.itemR[ri];
    }

    private static void add(List<Layer> out, TableStyle s, TableStyleType ty, boolean top, boolean right, boolean bottom, boolean left) {
        DifferentialStyleProvider p = s.getStyle(ty);
        if (p != null) out.add(new Layer(p, top, right, bottom, left, rank(ty)));
    }

    /** 要素の重なりの順位 (大きいほど上)。テーブルとピボットの順を 1 列に並べたもの。 */
    private static int rank(TableStyleType ty) {
        return switch (ty) {
            case wholeTable -> 0;
            case pageFieldLabels, pageFieldValues -> 1;
            case firstColumnStripe, secondColumnStripe -> 2;
            case firstRowStripe, secondRowStripe -> 3;
            case firstColumn -> 4;
            case lastColumn -> 5;
            case headerRow -> 6;
            case totalRow -> 7;
            case firstHeaderCell, lastHeaderCell, firstTotalCell, lastTotalCell -> 8;
            case firstSubtotalColumn, secondSubtotalColumn, thirdSubtotalColumn -> 9;
            case blankRow -> 10;
            case firstSubtotalRow, secondSubtotalRow, thirdSubtotalRow -> 11;
            case firstColumnSubheading, secondColumnSubheading, thirdColumnSubheading -> 12;
            case firstRowSubheading, secondRowSubheading, thirdRowSubheading -> 13;
        };
    }

    /** セルに効く塗りと文字。テーブルの外なら null。 */
    Look look(int r, int c) {
        Table t = any() ? at(r, c) : null;
        if (t == null) return null;
        return looks.computeIfAbsent((long) r * 16384 + c, k -> {
            Look l = new Look();
            for (Layer y : layers(t, r, c)) {
                PatternFormatting pf = y.p.getPatternFormatting();
                if (pf != null) {
                    Color f = color(pf.getFillBackgroundColorColor());
                    if (f == null) f = color(pf.getFillForegroundColorColor());
                    if (f != null) l.fill = f;
                }
                FontFormatting ff = y.p.getFontFormatting();
                if (ff != null) {
                    Color f = color(ff.getFontColor());
                    if (f != null) l.font = f;
                    CTFont ct = ctFont(ff);
                    if (ct != null) {
                        if (ct.sizeOfBArray() > 0) l.bold = !ct.getBArray(0).isSetVal() || ct.getBArray(0).getVal();
                        if (ct.sizeOfIArray() > 0) l.italic = !ct.getIArray(0).isSetVal() || ct.getIArray(0).getVal();
                        if (ct.sizeOfStrikeArray() > 0) l.strike = !ct.getStrikeArray(0).isSetVal() || ct.getStrikeArray(0).getVal();
                        if (ct.sizeOfUArray() > 0) {
                            var u = ct.getUArray(0);
                            String v = u.isSetVal() ? u.getVal().toString() : "single";
                            l.underline = v.equals("none") ? 0 : v.startsWith("double") ? 2 : 1;
                        }
                    } else {
                        if (ff.isBold()) l.bold = true;
                        if (ff.isItalic()) l.italic = true;
                    }
                }
            }
            if (t.pivot && c < t.dataC0 && r >= t.bodyR0 && r <= t.r1) {
                int kind = t.rowKind[r - t.bodyR0];
                if ((kind == ITEM || kind == SUBHEAD) && t.button[c - t.c0]) l.pivotLabel = true;
            }
            return l;
        });
    }

    /** dxf のフォント (明示の val="0" を読むため)。読めなければ null。 */
    private static CTFont ctFont(FontFormatting ff) {
        if (!(ff instanceof org.apache.poi.xssf.usermodel.XSSFFontFormatting)) return null;
        try {
            java.lang.reflect.Field f = org.apache.poi.xssf.usermodel.XSSFFontFormatting.class.getDeclaredField("_font");
            f.setAccessible(true);
            return (CTFont) f.get(ff);
        } catch (ReflectiveOperationException | RuntimeException e) {
            return null;
        }
    }

    /**
     * セルの 1 辺 (0 上、1 右、2 下、3 左) のテーブルの罫線 (なければ null)。優先順位の高い要素から探す。
     * 隣のセルも同じ表なら、境界の線は両方のセルの要素のうち順位の高いほうの線 (ピボットの見出しの下線が、本体の 1 行目の
     * テーブル全体の横線より優先する)。
     */
    Edge edge(int r, int c, int side, boolean black) {
        Table t = any() ? at(r, c) : null;
        if (t == null) return null;
        Object[] a = edgeRanked(t, r, c, side, black);
        int nr = r + (side == 0 ? -1 : side == 2 ? 1 : 0), nc = c + (side == 3 ? -1 : side == 1 ? 1 : 0);
        if (nr >= 0 && nc >= 0 && at(nr, nc) == t) {
            Object[] b = edgeRanked(t, nr, nc, (side + 2) % 4, black);
            if (b != null && (a == null || (int) b[1] > (int) a[1])) return (Edge) b[0];
        }
        return a == null ? null : (Edge) a[0];
    }

    /** セルの 1 辺の罫線と、それを決めた要素の順位 (重ねた順。なければ null)。 */
    private Object[] edgeRanked(Table t, int r, int c, int side, boolean black) {
        List<Layer> ls = layers(t, r, c);
        for (int i = ls.size() - 1; i >= 0; i--) {
            Layer y = ls.get(i);
            BorderFormatting b = y.p.getBorderFormatting();
            if (b == null) continue;
            BorderStyle bs;
            org.apache.poi.ss.usermodel.Color col;
            switch (side) {
                case 0 -> { bs = y.top ? b.getBorderTop() : b.getBorderHorizontal(); col = y.top ? b.getTopBorderColorColor() : b.getHorizontalBorderColorColor(); }
                case 1 -> { bs = y.right ? b.getBorderRight() : b.getBorderVertical(); col = y.right ? b.getRightBorderColorColor() : b.getVerticalBorderColorColor(); }
                case 2 -> { bs = y.bottom ? b.getBorderBottom() : b.getBorderHorizontal(); col = y.bottom ? b.getBottomBorderColorColor() : b.getHorizontalBorderColorColor(); }
                default -> { bs = y.left ? b.getBorderLeft() : b.getBorderVertical(); col = y.left ? b.getLeftBorderColorColor() : b.getVerticalBorderColorColor(); }
            }
            if (bs == null || bs == BorderStyle.NONE) continue;
            Color cc = color(col);
            return new Object[] {new Edge(bs, cc == null || black ? Color.BLACK : cc, true), y.rank};
        }
        return null;
    }

    private Color color(org.apache.poi.ss.usermodel.Color c) {
        return c instanceof XSSFColor x ? theme.color(x) : null;
    }
}
