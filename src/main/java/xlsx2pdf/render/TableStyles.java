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

import xlsx2pdf.render.CondFormat.Edge;

/**
 * シートのテーブル (挿入 → テーブル) のスタイルを、セルごとの塗り・文字・罫線にする。定義は組み込み (POI の presetTableStyles.xml)
 * とブックの独自スタイル。要素の優先順位は低い順に テーブル全体、列の縞、行の縞、最初・最後の列、見出し・集計行。
 * セルの書式はテーブルより優先する (呼び出し側で扱う)。縞は見えている行・列で数える。
 */
final class TableStyles {
    /** セルに効くテーブルの塗りと文字 (なければ null、false)。 */
    static final class Look {
        Color fill, font;
        boolean bold, italic;
    }

    /** テーブルの 1 要素と、その要素の範囲の中でのセルの位置 (外周の辺か)。 */
    private record Layer(DifferentialStyleProvider p, boolean top, boolean right, boolean bottom, boolean left) {}

    /** 縞の 1 行 (1 列)。kind は 0 なし、1 第 1 の縞、2 第 2 の縞。first・last は縞の端か。 */
    private record Band(int kind, boolean first, boolean last) {}

    private static final class Table {
        int r0, c0, r1, c1, bodyR0, bodyR1;
        TableStyle style;
        TableStyleInfo info;
        Band[] rowBands, colBands;
    }

    private final List<Table> tables = new ArrayList<>();
    private final Map<Long, Look> looks = new HashMap<>();
    private final Theme theme;

    TableStyles(XSSFSheet sheet, SheetLayout L, Theme theme) {
        this.theme = theme;
        for (XSSFTable x : sheet.getTables()) {
            if (!styled(x)) continue;
            AreaReference a = x.getCellReferences();
            if (a == null) continue;
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
        for (Table t : tables) if (r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1) return t;
        return null;
    }

    /** 優先順位の低い順の要素。 */
    private List<Layer> layers(Table t, int r, int c) {
        List<Layer> out = new ArrayList<>(6);
        boolean tl = c == t.c0, tr = c == t.c1;
        add(out, t.style.getStyle(TableStyleType.wholeTable), r == t.r0, tr, r == t.r1, tl);
        boolean header = r < t.bodyR0, total = r > t.bodyR1;
        if (!header && !total) {
            if (t.info.isShowColumnStripes()) {
                Band b = t.colBands[c - t.c0];
                if (b != null) add(out, t.style.getStyle(b.kind == 1 ? TableStyleType.firstColumnStripe : TableStyleType.secondColumnStripe),
                        r == t.bodyR0, b.last, r == t.bodyR1, b.first);
            }
            if (t.info.isShowRowStripes()) {
                Band b = t.rowBands[r - t.bodyR0];
                if (b != null) add(out, t.style.getStyle(b.kind == 1 ? TableStyleType.firstRowStripe : TableStyleType.secondRowStripe),
                        b.first, tr, b.last, tl);
            }
        }
        if (t.info.isShowFirstColumn() && tl) add(out, t.style.getStyle(TableStyleType.firstColumn), r == t.r0, true, r == t.r1, true);
        if (t.info.isShowLastColumn() && tr) add(out, t.style.getStyle(TableStyleType.lastColumn), r == t.r0, true, r == t.r1, true);
        if (header) add(out, t.style.getStyle(TableStyleType.headerRow), r == t.r0, tr, r == t.bodyR0 - 1, tl);
        if (total) add(out, t.style.getStyle(TableStyleType.totalRow), r == t.bodyR1 + 1, tr, r == t.r1, tl);
        return out;
    }

    private static void add(List<Layer> out, DifferentialStyleProvider p, boolean top, boolean right, boolean bottom, boolean left) {
        if (p != null) out.add(new Layer(p, top, right, bottom, left));
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
                    if (ff.isBold()) l.bold = true;
                    if (ff.isItalic()) l.italic = true;
                }
            }
            return l;
        });
    }

    /** セルの 1 辺 (0 上、1 右、2 下、3 左) のテーブルの罫線 (なければ null)。優先順位の高い要素から探す。 */
    Edge edge(int r, int c, int side, boolean black) {
        Table t = any() ? at(r, c) : null;
        if (t == null) return null;
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
            return new Edge(bs, cc == null || black ? Color.BLACK : cc);
        }
        return null;
    }

    private Color color(org.apache.poi.ss.usermodel.Color c) {
        return c instanceof XSSFColor x ? theme.color(x) : null;
    }
}
