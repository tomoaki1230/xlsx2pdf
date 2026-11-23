package xlsx2pdf.render;

import java.awt.Color;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellAlignment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;

import xlsx2pdf.render.CondFormat.Edge;

/** セルのブロックの背景・罫線・文字を描く。 */
public final class CellPainter {
    /** 横の余白: 用紙上で一定 (1.5pt) なので、縮小するとシートの単位では大きくなる。 */
    static final double PAD_PRINT = 1.5;
    /** 上・下の余白。 */
    static final double PAD_TOP = 0.6;
    static final double PAD_BOTTOM = 0.84;
    /** 既定の動きの倍率 100% の左右の余白 (pt)。実測に合わせた値。 */
    static final double POSM_PAD100_MODERN = 1.9;
    /** 縮小したときの左右の余白の下限 (用紙の pt)。 */
    static final double POSM_PADMIN = 0.85;
    /** 中央揃えの文字の横の補正 (用紙の pt)。 */
    static final double POSM_CENTER = 0.45;
    /** 細線の太さ (pt)。 */
    static final float THIN = 0.96f;
    /** 折り返した行の行送りの倍率 (フォントの高さを使うとき)。 */
    static final double LINE_FACTOR = 1.0;
    /** 字下げを用紙上で一定にする倍率の下限。 */
    static final double INDENT_PAPER_MIN_SCALE = 0.5;
    /** Excel 2013 以前で保存したブック (AppVersion ≤ 15) の文字のベースラインの下げ幅 (pt)。 */
    static final double LEGACY_TEXT_DY = 0.5;

    private final RenderContext ctx;
    private final SheetLayout L;
    private final XSSFSheet sheet;
    /** シートのテーブルのスタイル。 */
    private final TableStyles tables;
    /** ランを作っているセルのテーブルの文字の書式 (なければ null)。 */
    private TableStyles.Look runLook;
    private final PrintSetup ps;
    private final CondFormat cf;
    /** はみ出しを調べるだけで描かない段階か ({@link #paint})。 */
    private boolean measuringOverflow;
    /** {@link #overflowLastColumn} の計算中か、その結果 (文字が届く最後の列)。 */
    private boolean measuringRange;
    private int rangeLast;
    /** はみ出した文字が横切る縦罫線 (行 × 100000 + その左端の列)。 */
    private final java.util.Set<Long> overflowEdges = new java.util.HashSet<>();
    /** 列の書式 (列ごと。なければ null)。 */
    private final XSSFCellStyle[] colStyles;

    /** 記録用: ページ番号と、ブックの座標 (シートの pt) からページの座標への変換 (page = o + (sheet - s) × scale)。 */
    private double blockOx, blockOy, blockSx, blockSy;
    /** 描いているセルの均等割り付けで、前後にスペースを入れるか (justifyLastLine)。 */
    private boolean spaceEnds;
    /** 記録用: 描いているセルの番地と配置。 */
    /** 描いているセルが縦書きか。 */
    private boolean stackedText;
    /** 描いているセルが結合セルか。 */
    private boolean mergedText;
    /** 描いているセルが書式付き文字列か。 */
    private boolean cellRich;
    /** 描いているセルが折り返しの設定か。 */
    private boolean cellWrap;
    /** 折り返して配置するか (折り返しの設定に加え、横の両端揃え・均等割り付けなども折り返す)。 */
    private boolean cellWrapLayout;

    /** これから描くブロックのページ上の配置 (ox、oy は用紙上の左上、sx、sy はそのシートの座標)。 */
    public void setBlockPlacement(double ox, double oy, double sx, double sy) {
        blockOx = ox; blockOy = oy; blockSx = sx; blockSy = sy;
    }

    /**
     * セルの描画器を作る。
     *
     * @param cf 条件付き書式
     */
    public CellPainter(RenderContext ctx, SheetLayout L, PrintSetup ps, CondFormat cf) {
        this.ctx = ctx;
        this.L = L;
        this.sheet = L.sheet;
        this.tables = new TableStyles(L.sheet, L, ctx.theme);
        this.ps = ps;
        this.cf = cf;
        colStyles = new XSSFCellStyle[L.maxCol];
        for (CTCols cols : sheet.getCTWorksheet().getColsArray()) {
            for (CTCol c : cols.getColArray()) {
                if (!c.isSetStyle() || c.getStyle() == 0) continue;
                XSSFCellStyle st = ctx.wb.getCellStyleAt((int) c.getStyle());
                for (long i = Math.max(1, c.getMin()) - 1; i < Math.min(c.getMax(), L.maxCol); i++) colStyles[(int) i] = st;   // min="0" などの不正な値で負の添字にしない
            }
        }
    }

    private double sheetDescentPt = -1;

    /** 下揃えの文字を置く行 (結合セルは最後の行)。{@link #rowDescent} に使う。 */
    private int textBottomRow = -1;
    private final java.util.Map<Integer, Double> rowDescentPt = new java.util.HashMap<>();

    /**
     * 行の下揃えのセル (文字のない書式だけのセルも含む。縦の配置が下・指定なしで回転なし) のフォントのディセントの最大 (pt)。
     * Excel は行の下揃えの文字を共通の基線に置き、その基線は行の中で一番ディセントの大きい書体で決まる (例: 同じ行に
     * メイリオ 11 の下揃えのセルがあると、ＭＳ Ｐゴシック 11 の文字は約 4pt 上がる。中央・上揃えのセルは関係しない)。
     */
    private double descentPt(XSSFFont f) {
        return ctx.fonts.spec(f.getFontName(), false, false).descent * (f.getFontHeight() / 20.0);
    }

    private double rowDescent(int r) {
        if (r < 0) return 0;
        Double v = rowDescentPt.get(r);
        if (v != null) return v;
        double m = 0;
        var row = sheet.getRow(r);
        if (row != null) for (var cl : row) {
            if (!(cl.getCellStyle() instanceof XSSFCellStyle st)) continue;
            Align a = align(st);
            if (a.rotation() != 0 || (a.v() != VerticalAlignment.BOTTOM && a.v() != null)) continue;
            XSSFFont f = st.getFont();
            if (f == null) continue;
            m = Math.max(m, descentPt(f));
        }
        rowDescentPt.put(r, m);
        return m;
    }

    /** シートの標準行を決めるフォントのディセントの最大 (pt)。 */

    /** 画面 (96 dpi) の整数の画素に丸めたディセント (pt): round(ディセント (em) × round(大きさ × 4/3)) × 0.75。 */
    private double screenDescPt(String font, double size) {
        return Math.round(ctx.fonts.spec(font, false, false).descent * Math.round(size * 4 / 3.0)) * 0.75;
    }

    /** 画面の 1 画素の差を、プリンタの何画素として数えるか (6.25 ではなく約 7.25。14〜24pt で実測)。 */
    static final double SCR_STEP = 7.25 / 6.25;

    /** 行の標準のフォントのディセント (pt)。書式のある行 (row の customFormat と s) で書式のフォントが日本語の名前の書体ならそのフォント、ほかはシートの標準行のフォント。 */
    private double baseDescent(int r) {
        var row = r >= 0 ? sheet.getRow(r) : null;
        if (row != null && row.getCTRow().isSetCustomFormat() && row.getCTRow().getCustomFormat() && row.getCTRow().isSetS()) {
            long si = row.getCTRow().getS();
            if (si >= 0 && si < sheet.getWorkbook().getNumCellStyles()) {
                XSSFFont f = sheet.getWorkbook().getCellStyleAt((int) si).getFont();
                // 欧文の書体 (Century など) の書式の行は、シートの標準行のフォントのまま
                if (f != null && f.getFontName() != null && !f.getFontName().chars().allMatch(ch -> ch < 0x80)) return descentPt(f);
            }
        }
        return sheetDescent();
    }

    private double sheetDescent() {
        if (sheetDescentPt < 0) {
            double m = 0;
            for (String k : L.sheetRowFonts()) {
                int bar = k.lastIndexOf('|');
                m = Math.max(m, ctx.fonts.spec(k.substring(0, bar), false, false).descent * Double.parseDouble(k.substring(bar + 1)));
            }
            sheetDescentPt = m;
        }
        return sheetDescentPt;
    }

    /** 描いているセルのフォント (横の余白の計算に使う)。 */
    private XSSFFont padFont;

    /** 既定の動きの位置の式を使うか。 */
    private boolean posx() { return newLayout() && !L.model().authorEnv; }

    /** 罫線の中心のセルの境界からのずれ (プリンタの画素)。 */
    private double lineOff() { return Math.floor(8 * L.scale + 0.5 + 1e-9) % 2 == 1 ? 4.5 : 4; }

    /** 上の余白 (シートの pt)。基線 = 行の上端 + 10 画素 + アセント。 */
    double padTop() { return posx() ? (10 - lineOff()) * 0.12 / Math.max(0.1, L.scale) : PAD_TOP; }

    /** 下の余白 (シートの pt)。基線 = 行の下端 + 8 − 9.5 × 倍率 画素 − ディセント。 */
    double padBottom() { return posx() ? (9.5 * L.scale - 8 + lineOff()) * 0.12 / Math.max(0.1, L.scale) : PAD_BOTTOM; }

    /** 中央揃えの縦のずれ (シートの pt)。行の中央の計算値 + 5 画素。 */
    /**
     * 文字の送り幅を測る格子 (シートの pt)。線形に縮小したシートも、文字は倍率で実体化した幅で並ぶ (10pt × 77%: ppem 64 の全角 64 画素。100% の 84 画素 × 77% ではない)。
     */
    private double textGrid() { return !L.model().authorEnv ? 0.12 / Math.max(0.01, L.scale) : L.devUnitX(); }

    /** 結合セルの縦の中央揃えの文字を、結合していないセルより上へ動かす量 (プリンタの画素)。 */
    static final double MERGE_CENTER_PX = 4.0;

    /** 縦書きの中央揃えの文字を、セルの中央より下へ動かす量 (プリンタの画素)。 */
    static final double STACK_CENTER_PX = 5.0;

    /** 結合セルの折り返しの下揃えの文字を、結合していないセルより上へ動かす最小の量 (プリンタの画素)。 */
    static final double MERGE_WRAP_BOTTOM_PX = 6.0;

    double centerDy() { return posx() ? (5 - lineOff()) * 0.12 / Math.max(0.1, L.scale) : 0; }

    /** セルの境界から文字までの余白 (シートの pt): ceil(セルのフォントのプリンタの数字幅 / 4) 画素。 */
    private double digitPad() {
        double sc = Math.max(0.1, L.scale);
        double size = !Double.isNaN(padSize) ? padSize : padFont != null ? padFont.getFontHeight() / 20.0 : L.model().defaultSize();
        double[] g = padFont != null ? FontManager.gdiMetrics(ctx.cellFontName(padFont)) : null;
        double em = g != null ? g[2] : L.model().digitEm;
        long ppem = (long) Math.floor(Math.floor(size * 20 * sc + 0.5 + 1e-9) * 5 / 12 + 0.5 + 1e-9);
        long digit = Math.round(em * ppem);
        // 余白は文字の大きさで決まる: ceil(P / 8) 画素
        return Math.ceil(digit / 4.0) * 0.12 / sc;
    }

    /** 余白の計算に使うフォントの大きさ (pt)。NaN ならセルのフォントの大きさ。 */
    private double padSize = Double.NaN;

    /**
     * 縮小して全体を表示: プリンタ (600 dpi) の ppem を元の大きさから 1 ずつ下げ、文字の幅 (格子に吸着した送り幅) + 縮めた後のフォントの
     * 左右の余白がセルの幅 w に収まる最初の ppem にする。描く大きさは ppem × 0.12pt。複数のランは同じ比で ppem を下げる。
     * 始める ppem は元の大きさでの比 (floor(元の ppem × (セルの幅 − 元の余白) ÷ 元の文字の幅))。Excel と同じ大きさになるのは約 8 割。
     *
     * @param extraL 左の余白に足す字下げ (pt)
     * @param extraR 右の余白に足す字下げ (pt)
     * @return {左の余白, 右の余白}
     */
    private double[] shrinkSteps(List<TextLayout.Run> runs, double w, double extraL, double extraR) {
        double sc = Math.max(0.1, L.scale);
        float[] base = new float[runs.size()];
        for (int i = 0; i < base.length; i++) base[i] = runs.get(i).size;
        float first = base[0];
        long ppem0 = ppemOf(first, sc);
        double padL = padX() + extraL, padR = padX() + extraR;
        int k0 = 1;
        {
            // Excel は元の大きさでの比で ppem を決める: floor(元の ppem × (セルの幅 − 元の左右の余白) ÷ 元の文字の幅)。
            // その大きさで (縮めた後の余白を含めて) 収まらなければ、収まるまで 1 ずつ下げる
            double tw0 = totalWidth(runs);
            if (tw0 > 0) {
                long pr = (long) Math.floor(ppem0 * (w - padL - padR) / tw0 + 1e-9);
                k0 = (int) Math.max(1, Math.min(ppem0 - 1, ppem0 - pr));   // 使える幅が 0 以下でも 1 ppem までは縮める
            }
        }
        for (int k = k0; k < 2000; k++) {
            long p0 = ppem0 - k;
            if (p0 < 1) break;
            double ratio = (double) p0 / ppem0;
            for (int i = 0; i < base.length; i++) {
                long p = Math.round(ppemOf(base[i], sc) * ratio);
                runs.get(i).size = (float) (Math.max(1, p) * 0.12 / sc);
            }
            padSize = runs.get(0).size;
            padL = padX() + extraL;
            padR = padX() + extraR;
            padSize = Double.NaN;
            if (totalWidth(runs) + padL + padR <= w + 1e-6) break;
        }
        return new double[] {padL, padR};
    }

    /** フォントの大きさ (pt) の 600 dpi のプリンタでの ppem。 */
    private static long ppemOf(double size, double sc) {
        return (long) Math.floor(Math.floor(size * 20 * sc + 0.5 + 1e-9) * 5 / 12 + 0.5 + 1e-9);
    }

    /** 横の余白 (シートの pt)。 */
    double padX() {
        if (newLayout() && !L.model().authorEnv) {
            // 罫線の中心から 4 + ceil(セルのフォントのプリンタの数字幅 / 4) 画素
            return 4 * 0.12 / Math.max(0.1, L.scale) + digitPad();
        }
        if (newLayout()) return Math.max(POSM_PADMIN / Math.max(0.1, L.scale), POSM_PAD100_MODERN);
        return PAD_PRINT / Math.max(0.1, L.scale);
    }

    /** 新しい Excel の配置 (線形に縮小するシートか倍率 100%) か。 */
    private boolean newLayout() {
        return L.linear || L.scale == 1;
    }

    /**
     * セルの配置。
     *
     * @param h        横位置
     * @param v        縦位置
     * @param wrap     折り返して全体を表示する
     * @param rotation 文字の角度 (0〜180。255 は縦書き)
     * @param indent   字下げ
     * @param shrink   縮小して全体を表示する
     * @param spaceEnds 均等割り付けで前後にスペースを入れる (justifyLastLine)
     */
    record Align(HorizontalAlignment h, VerticalAlignment v, boolean wrap, int rotation, int indent, boolean shrink, boolean spaceEnds) {}

    /** 書式の配置を読む ({@link #ALIGN_MODE})。 */
    static Align align(XSSFCellStyle st) {
        CTXf xf = st.getCoreXf();
        CTCellAlignment a = xf != null && xf.isSetAlignment() ? xf.getAlignment() : null;
        if (a == null) return new Align(HorizontalAlignment.GENERAL, VerticalAlignment.BOTTOM, false, 0, 0, false, false);
        HorizontalAlignment h = !a.isSetHorizontal() ? HorizontalAlignment.GENERAL : switch (a.getHorizontal().toString()) {
            case "left" -> HorizontalAlignment.LEFT;
            case "center" -> HorizontalAlignment.CENTER;
            case "right" -> HorizontalAlignment.RIGHT;
            case "fill" -> HorizontalAlignment.FILL;
            case "justify" -> HorizontalAlignment.JUSTIFY;
            case "centerContinuous" -> HorizontalAlignment.CENTER_SELECTION;
            case "distributed" -> HorizontalAlignment.DISTRIBUTED;
            default -> HorizontalAlignment.GENERAL;
        };
        VerticalAlignment v = !a.isSetVertical() ? VerticalAlignment.BOTTOM : switch (a.getVertical().toString()) {
            case "top" -> VerticalAlignment.TOP;
            case "center" -> VerticalAlignment.CENTER;
            case "justify" -> VerticalAlignment.JUSTIFY;
            case "distributed" -> VerticalAlignment.DISTRIBUTED;
            default -> VerticalAlignment.BOTTOM;
        };
        return new Align(h, v, a.isSetWrapText() && a.getWrapText(), a.isSetTextRotation() ? a.getTextRotation().intValue() : 0,
                a.isSetIndent() ? (int) a.getIndent() : 0, a.isSetShrinkToFit() && a.getShrinkToFit(),
                a.isSetJustifyLastLine() && a.getJustifyLastLine());
    }

    /** セル (なければ null)。 */
    XSSFCell cell(int r, int c) {
        XSSFRow row = sheet.getRow(r);
        return row == null ? null : row.getCell(c);
    }

    /** セルの書式。セルがなければ行の書式、それもなければ列の書式 (なければ null)。 */
    XSSFCellStyle style(int r, int c) {
        XSSFRow row = sheet.getRow(r);
        if (row != null) {
            XSSFCell cell = row.getCell(c);
            if (cell != null) return cell.getCellStyle();
            if (row.isFormatted()) return row.getRowStyle();
        }
        return c < colStyles.length ? colStyles[c] : null;
    }

    /** セルが空 (なし、空白、または空文字列) か。 */
    public static boolean isEmptyCell(XSSFCell cell) { return isEmpty(cell); }

    /** シートのセル (なければ null)。 */
    private XSSFCell cellAt(int r, int c) {
        var row = sheet.getRow(r);
        return row == null ? null : (XSSFCell) row.getCell(c);
    }

    /** セルが空か ({@link #isEmptyCell} と同じ)。 */
    static boolean isEmpty(XSSFCell cell) {
        if (cell == null) return true;
        CellType t = cell.getCellType();
        if (t == CellType.BLANK) return true;
        if (t == CellType.STRING) return cell.getStringCellValue().isEmpty();
        return false;
    }

    /**
     * 文字のはみ出しまで含めた最後の列を求める。印刷範囲のないシートでは、Excel はセルの文字がはみ出した列まで印刷する。判定はプリンタの画素で、
     * 列の幅の合計が「文字の幅 + ceil(数字幅 / 4)」以上になる列まで。使用範囲の外まで届くのは、行の一番右の文字だけ。
     *
     * @param lastRow 最後の行
     * @param lastCol 最後の列
     * @param limit   調べる列の上限
     * @return はみ出しまで含めた最後の列
     */
    public int overflowLastColumn(int lastRow, int lastCol, int limit) {
        int last = lastCol;
        measuringOverflow = true;
        measuringRange = true;
        try {
            for (var row : sheet) {
                int r = row.getRowNum();
                if (r > lastRow) break;
                if (L.rh(r) <= 0) continue;
                // 右へ傾けた文字のセル: 平行四辺形の右上の角がかかる列まで印刷する
                for (var cl : row) {
                    if (!(cl.getCellStyle() instanceof XSSFCellStyle cst)) continue;
                    double sl = slantAngle(cst);
                    int cc = cl.getColumnIndex();
                    if (sl <= 0 || L.cw(cc) <= 0 || L.mergedAt(r, cc) != null || isEmpty((XSSFCell) cl)) continue;
                    rangeLast = cc;
                    try {
                        paintCellText(null, new Paginator.Segment(r, r, cc, limit, 0, 0, true), (XSSFCell) cl, null);
                    } catch (IOException | RuntimeException e) {
                        continue;
                    }
                    last = Math.max(last, rangeLast);
                }
                int c = -1;
                for (int k = row.getLastCellNum() - 1; k >= 0 && c < 0; k--) {
                    // 非表示の列の文字は印刷されないので、範囲を広げない
                    if (L.cw(k) > 0 && !isEmpty((XSSFCell) row.getCell(k))) c = k;
                }
                if (c < 0 || L.mergedAt(r, c) != null) continue;
                rangeLast = c;
                try {
                    paintTexts0(null, new Paginator.Segment(r, r, c, limit, 0, 0, true));
                } catch (IOException e) {
                    continue;   // 測れないセルは範囲を広げない
                }
                last = Math.max(last, rangeLast);
            }
        } finally {
            measuringOverflow = false;
            measuringRange = false;
        }
        return last;
    }

    /** プリンタの画素で、文字が届く最後の列を求める ({@link #overflowLastColumn})。 */
    private void measureRange(int r, int c, Paginator.Segment s, HorizontalAlignment ha, List<TextLayout.Run> runs, double padL) {
        // 描画中 (paint) から呼ばれることがあるので、終わったら元の格子に戻す
        double prev = FontManager.gridUnit();
        FontManager.setGrid(textGrid());
        double need;
        try {
            need = totalWidth(runs) + padL - padX() + digitPad();
        } finally {
            FontManager.setGrid(prev);
        }
        double w = L.cw(c);
        if (ha == HorizontalAlignment.CENTER) need = w + (need - w) / 2;
        int cc = c;
        while (w < need - 0.001 && cc < s.c1 && isEmpty(cell(r, cc + 1)) && L.mergedAt(r, cc + 1) == null) {
            cc++;
            w += L.cw(cc);
        }
        rangeLast = cc;
    }

    /** ブロックのすべての層 (塗りつぶし、枠線、罫線、文字) を描く (座標はシートの pt)。 */
    public void paint(PdfCanvas cv, Paginator.Segment s) throws IOException {
        if (!ps.draft) paintFills(cv, s);
        if (ps.gridlines && !ps.draft) paintGridlines(cv, s);
        // Excel はセルの文字を GDI でプリンタのピクセル単位にレイアウトする: 各グリフの送り幅は整数デバイスピクセル
        FontManager.setGrid(textGrid());
        try {
            // 文字がはみ出す範囲を先に調べる (はみ出した文字が横切る縦罫線は描かない)
            overflowEdges.clear();
            measuringOverflow = true;
            try {
                paintTexts(cv, s);
            } finally {
                measuringOverflow = false;
            }
            if (!ps.draft) paintBorders(cv, s);
            paintTexts(cv, s);
        } finally {
            FontManager.clearGrid();
        }
        if (!ps.draft) paintSparklines(cv, s);
    }

    /** シートのスパークライン (遅延して読む。なければ null)。 */
    private Sparklines sparklines;
    private boolean sparkRead;

    /** ブロックの中のセルのスパークラインを描く。 */
    private void paintSparklines(PdfCanvas cv, Paginator.Segment s) throws IOException {
        if (!sparkRead) {
            sparkRead = true;
            try {
                sparklines = Sparklines.parse(ctx, L.sheet);
            } catch (RuntimeException e) {
                Warnings.add("スパークラインを読めませんでした", e);
            }
        }
        if (sparklines == null) return;
        for (int[] rc : sparklines.cells()) {
            int r = rc[0], c = rc[1];
            if (r < s.r0 || r > s.r1 || c < s.c0 || c > s.c1) continue;
            double w = L.cw(c), h = L.rh(r);
            if (w <= 0 || h <= 0) continue;
            CellRangeAddress m = L.mergedAt(r, c);
            double x = L.x(c), y = L.y(r);
            if (m != null) { x = L.x(m.getFirstColumn()); y = L.y(m.getFirstRow()); w = L.x(m.getLastColumn() + 1) - x; h = L.y(m.getLastRow() + 1) - y; }
            int depth = cv.depth();
            try {
                sparklines.paint(cv, r, c, x, y, w, h);
            } catch (RuntimeException e) {
                cv.restoreTo(depth);
                Warnings.add("スパークラインを描けませんでした", e);
            }
        }
    }

    // -------------------------------------------------------------- 塗りつぶし

    /** セルの塗りつぶし (条件付き書式・グラデーション・データバーを含む) を描く。 */
    private void paintFills(PdfCanvas cv, Paginator.Segment s) throws IOException {
        // 白黒印刷 (pageSetup blackAndWhite) ではセルの塗りつぶしを印刷しない
        if (bw()) return;
        for (int r = s.r0; r <= s.r1; r++) {
            ctx.checkDeadline();
            if (L.rh(r) <= 0) continue;
            for (int c = s.c0; c <= s.c1; c++) {
                if (L.cw(c) <= 0) continue;
                CellRangeAddress m = L.mergedAt(r, c);
                int rr = r, cc = c;
                double x, y, w, h;
                if (m != null) {
                    // 結合範囲は、このブロック内で最初に見えるセルで 1 回だけ描く
                    if (r != firstVisibleRow(m, s) || c != firstVisibleColumn(m, s)) continue;
                    rr = m.getFirstRow();
                    cc = m.getFirstColumn();
                    x = L.x(m.getFirstColumn()); y = L.y(m.getFirstRow());
                    w = L.x(m.getLastColumn() + 1) - x; h = L.y(m.getLastRow() + 1) - y;
                } else {
                    x = L.x(c); y = L.y(r); w = L.cw(c); h = L.rh(r);
                }
                XSSFCellStyle st = style(rr, cc);
                CondFormat.Result cfr = cf.eval(rr, cc);
                Color fill = cfr != null && cfr.fill != null ? cfr.fill : null;
                if (fill == null && cfr != null && cfr.patBits != 0) {
                    // 条件付き書式の網掛け
                    cv.fillPatternTile(x, y, w, h, cfr.patBits, cfr.patFg, cfr.patBg);
                    if (cfr.bar != null) paintDataBar(cv, cfr, x, y, w, h);
                    continue;
                }
                if (fill == null && st != null) {
                    if (paintGradient(cv, st, x, y, w, h)) {
                        if (cfr != null && cfr.bar != null) paintDataBar(cv, cfr, x, y, w, h);
                        continue;
                    }
                    if (paintPattern(cv, st, x, y, w, h)) {
                        if (cfr != null && cfr.bar != null) paintDataBar(cv, cfr, x, y, w, h);
                        continue;
                    }
                    fill = fillColor(st);
                }
                if (fill == null) {
                    TableStyles.Look lk = tables.look(rr, cc);
                    if (lk != null) fill = lk.fill;
                }
                // Excel は塗りの矩形を右と下に 1 画素 (600 dpi で 0.12pt) 広く塗る (次のセルの塗りに 1 画素重なる)
                if (fill != null) cv.fillRect(x, y, w + FILL_OVER / L.scale, h + FILL_OVER / L.scale, fill);
                if (cfr != null && cfr.bar != null) paintDataBar(cv, cfr, x, y, w, h);
            }
        }
    }

    /** セルの単色の塗りを右と下に広げる幅 (用紙の pt。600 dpi の 1 画素)。 */
    static final double FILL_OVER = 0.12;

    /** 書式の単色の塗りつぶしの色 (網掛けとなしは null)。 */
    Color fillColor(XSSFCellStyle st) {
        FillPatternType p = st.getFillPattern();
        if (p == null || p == FillPatternType.NO_FILL) return null;
        org.apache.poi.xssf.usermodel.XSSFColor fgx = st.getFillForegroundXSSFColor();
        // 単色の塗りの前景色が「自動」なら白
        if (p == FillPatternType.SOLID_FOREGROUND && fgx != null && fgx.isAuto()) return Color.WHITE;
        Color fg = fillPart(fgx);
        Color bg = fillPart(st.getFillBackgroundXSSFColor());
        if (p == FillPatternType.SOLID_FOREGROUND) return fg != null ? fg : (bg != null ? bg : Color.BLACK);
        // 網掛けは paintPattern で模様として描く
        return null;
    }

    /**
     * Excel が印刷する網掛けの模様 (600 dpi の 8 × 8 画素。ビット r × 8 + c が 1 なら前景色)。Excel の PDF から取り出した。
     * 色や倍率によらず同じ模様で、周期は用紙の座標に揃う。
     */
    static final java.util.Map<String, Long> PATTERNS = java.util.Map.ofEntries(
            Map.entry("gray125", 0x0010000100100001L),
            Map.entry("gray0625", 0x0000001000000001L),
            Map.entry("darkGray", 0xAA55AA55AA55AA55L),
            Map.entry("mediumGray", 0x4411441144114411L),
            Map.entry("lightGray", 0x0044001100440011L),
            Map.entry("darkHorizontal", 0x00000000FFFFFFFFL),
            Map.entry("darkVertical", 0x0F0F0F0F0F0F0F0FL),
            Map.entry("darkDown", 0x87C3E1F0783C1E0FL),
            Map.entry("darkUp", 0x78F0E1C3870F1E3CL),
            Map.entry("darkGrid", 0xF0F0F0F00F0F0F0FL),
            Map.entry("darkTrellis", 0xE7C3E7FF7E3C7EFFL),
            Map.entry("lightHorizontal", 0x000000000000FFFFL),
            Map.entry("lightVertical", 0x0303030303030303L),
            Map.entry("lightDown", 0x81C06030180C0603L),
            Map.entry("lightUp", 0x3060C08103060C18L),
            Map.entry("lightGrid", 0x030303030303FFFFL),
            Map.entry("lightTrellis", 0xB1E0E0B11B0E0E1BL));

    /** SpreadsheetML の patternType の名前 (POI の FillPatternType の番号の順)。 */
    static final String[] PATTERN_NAMES = {"none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal",
            "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown",
            "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"};

    /** 網掛けの塗りを、Excel と同じ模様で描く。網掛けでなければ何もせず false。 */
    private boolean paintPattern(PdfCanvas cv, XSSFCellStyle st, double x, double y, double w, double h) throws IOException {
        FillPatternType p = st.getFillPattern();
        if (p == null || p.getCode() < 2 || p.getCode() >= PATTERN_NAMES.length) return false;
        Long bits = PATTERNS.get(PATTERN_NAMES[p.getCode()]);
        if (bits == null) return false;
        Color fg = fillPart(st.getFillForegroundXSSFColor());
        Color bg = fillPart(st.getFillBackgroundXSSFColor());
        cv.fillPatternTile(x, y, w, h, bits, fg != null ? fg : Color.BLACK, bg != null ? bg : Color.WHITE);
        return true;
    }

    /** 塗りつぶしの色の部分。システム色はウィンドウの背景 (65、白) とウィンドウの文字 (64、黒)。 */
    private Color fillPart(org.apache.poi.xssf.usermodel.XSSFColor c) {
        if (c != null && c.isIndexed() && !c.isRGB() && !c.isThemed()) {
            if (c.getIndex() == 65) return Color.WHITE;
            if (c.getIndex() == 64) return Color.BLACK;
        }
        return ctx.theme.color(c);
    }

    /** 色を混ぜる (a の割合 t)。 */
    static Color mix(Color a, Color b, double t) {
        return new Color((int) Math.round(a.getRed() * t + b.getRed() * (1 - t)),
                (int) Math.round(a.getGreen() * t + b.getGreen() * (1 - t)),
                (int) Math.round(a.getBlue() * t + b.getBlue() * (1 - t)));
    }

    /** グラデーションの塗りつぶしを帯で近似して描く。グラデーションでなければ false。 */
    private boolean paintGradient(PdfCanvas cv, XSSFCellStyle st, double x, double y, double w, double h) throws IOException {
        CTXf xf = st.getCoreXf();
        if (xf == null || !xf.isSetFillId()) return false;
        var fill = ctx.wb.getStylesSource().getFillAt((int) xf.getFillId()).getCTFill();
        if (fill == null || !fill.isSetGradientFill()) return false;
        CTGradientFill g = fill.getGradientFill();
        CTGradientStop[] stops = g.getStopArray();
        if (stops.length == 0) return false;
        List<double[]> pts = new ArrayList<>();
        List<Color> cols = new ArrayList<>();
        for (CTGradientStop sp : stops) {
            XSSFColor xc = XSSFColor.from(sp.getColor(), ctx.wb.getStylesSource().getIndexedColors());
            if (xc != null && xc.isThemed()) ctx.wb.getStylesSource().getTheme().inheritFromThemeAsRequired(xc);
            Color c = ctx.theme.color(xc);
            pts.add(new double[] {sp.getPosition()});
            cols.add(c == null ? Color.WHITE : c);
        }
        boolean path = g.isSetType() && "path".equals(g.getType().toString());
        double deg = g.isSetDegree() ? g.getDegree() : 0;
        cv.save();
        try {
            cv.clipRect(x, y, w, h);
            if (path) {
                // 内側の矩形 (left・right・top・bottom の割合。t = 0) からセルの縁 (t = 1) へ広がる同心の矩形。外側から塗る
                double il = g.isSetLeft() ? g.getLeft() : 0, ir = g.isSetRight() ? g.getRight() : 0;
                double it = g.isSetTop() ? g.getTop() : 0, ib = g.isSetBottom() ? g.getBottom() : 0;
                int n = PATH_BANDS;
                for (int i = n; i >= 1; i--) {
                    double t = i / (double) n;
                    Color c = gradAt(pts, cols, t - 0.5 / n);
                    double l = il * (1 - t), r = ir + (1 - ir) * t, tp = it * (1 - t), bt = ib + (1 - ib) * t;
                    cv.fillRect(x + w * l, y + h * tp, w * (r - l), h * (bt - tp), c);
                }
            } else if (Math.abs(deg - 90 * Math.round(deg / 90)) < 1e-9) {
                // 縦・横 (0・90・180・270 度): Excel の印刷と同じく、各チャンネルを 2 段ずつ変える帯を並べる。帯の数は区間ごとに
                // 最大のチャンネルの差 ÷ 2 (白 → 0070C0 で 128) と終わりの色の 1 本、幅は長さ ÷ 本数
                int q = (int) Math.floorMod(Math.round(deg / 90), 4);
                boolean horizontal = q % 2 == 0, reverse = q >= 2;
                List<Color> bands = new ArrayList<>();
                for (int i = 0; i + 1 < cols.size(); i++) {
                    Color a = cols.get(i), b = cols.get(i + 1);
                    int diff = Math.max(Math.abs(a.getRed() - b.getRed()), Math.max(Math.abs(a.getGreen() - b.getGreen()), Math.abs(a.getBlue() - b.getBlue())));
                    double segLen = pts.get(i + 1)[0] - pts.get(i)[0];
                    int n = Math.max(1, (int) Math.ceil(diff / 2.0));
                    if (segLen <= 0) continue;
                    for (int k = 0; k < n; k++) bands.add(lerp(a, b, k / (double) n));
                }
                // 最初の位置より前・最後の位置より後は端の色
                double p0 = pts.get(0)[0], p1 = pts.get(pts.size() - 1)[0];
                bands.add(cols.get(cols.size() - 1));
                double len = horizontal ? w : h;
                int nb = bands.size();
                double a0 = len * p0, a1 = len * p1, step = (a1 - a0) / Math.max(1, nb);
                if (a0 > 0) bandRect(cv, x, y, w, h, horizontal, reverse, 0, a0, len, cols.get(0));
                for (int k = 0; k < nb; k++) bandRect(cv, x, y, w, h, horizontal, reverse, a0 + k * step, a0 + (k + 1) * step, len, bands.get(k));
                if (a1 < len) bandRect(cv, x, y, w, h, horizontal, reverse, a1, len, len, cols.get(cols.size() - 1));
            } else {
                // 直線: セルを単位の正方形とみなした角度 (0 は左→右、90 は上→下) の方向に変わる。PDF のグラデーションで塗る
                double rad = Math.toRadians(deg), cs = Math.cos(rad), sn = Math.sin(rad), d = Math.abs(cs) + Math.abs(sn);
                cv.transform(new java.awt.geom.AffineTransform(w, 0, 0, h, x, y));
                double[] pos = new double[pts.size()];
                for (int i = 0; i < pos.length; i++) pos[i] = pts.get(i)[0];
                cv.fillAxialStops(0.5 - cs * d / 2, 0.5 - sn * d / 2, 0.5 + cs * d / 2, 0.5 + sn * d / 2, pos, cols.toArray(new Color[0]));
            }
        } finally {
            cv.restore();
        }
        return true;
    }

    /** 2 色の間の色 (t = 0 で a、1 で b)。 */
    private static Color lerp(Color a, Color b, double t) {
        return new Color((int) Math.round(a.getRed() + (b.getRed() - a.getRed()) * t), (int) Math.round(a.getGreen() + (b.getGreen() - a.getGreen()) * t),
                (int) Math.round(a.getBlue() + (b.getBlue() - a.getBlue()) * t));
    }

    /** 縦・横のグラデーションの帯 (軸の上の [from, to] pt)。帯の縁は 600dpi の画素 (0.12pt) に揃える。 */
    private static void bandRect(PdfCanvas cv, double x, double y, double w, double h, boolean horizontal, boolean reverse,
                                 double from, double to, double len, Color c) throws IOException {
        from = Math.round(from / 0.12) * 0.12; to = Math.min(len, Math.round(to / 0.12) * 0.12);
        if (to <= from) return;
        double f = reverse ? len - to : from, t = reverse ? len - from : to;
        if (horizontal) cv.fillRect(x + f, y, t - f, h, c); else cv.fillRect(x, y + f, w, t - f, c);
    }

    /** 中心から広がるグラデーションの帯の数。 */
    static final int PATH_BANDS = 64;

    /** セルの書式がグラデーションの塗りか (印刷の範囲に数える)。 */
    public static boolean hasGradient(org.apache.poi.ss.usermodel.Cell cell) {
        if (!(cell.getCellStyle() instanceof XSSFCellStyle st)) return false;
        CTXf xf = st.getCoreXf();
        if (xf == null || !xf.isSetFillId()) return false;
        try {
            var f = ((org.apache.poi.xssf.usermodel.XSSFWorkbook) cell.getSheet().getWorkbook()).getStylesSource().getFillAt((int) xf.getFillId()).getCTFill();
            return f != null && f.isSetGradientFill();
        } catch (RuntimeException e) {
            return false;
        }
    }

    /** グラデーションの位置 t の色。 */
    static Color gradAt(List<double[]> pos, List<Color> cols, double t) {
        if (t <= pos.get(0)[0]) return cols.get(0);
        for (int i = 1; i < pos.size(); i++) {
            if (t <= pos.get(i)[0]) {
                double a = pos.get(i - 1)[0], b = pos.get(i)[0];
                return mix(cols.get(i), cols.get(i - 1), b > a ? (t - a) / (b - a) : 1);
            }
        }
        return cols.get(cols.size() - 1);
    }

    /** データバーを描く。 */
    private void paintDataBar(PdfCanvas cv, CondFormat.Result r, double x, double y, double w, double h) throws IOException {
        DataBars.Bar b = r.bar;
        // 棒を描ける範囲: 始まりの側 1.92pt、終わりの側 0.96pt、上 1.92pt、下 2.24pt の内側 (Excel の印刷)
        double left = x + (b.rtl() ? DB_END : DB_START), len = w - DB_START - DB_END;
        double top = y + DB_TOP, bh = h - DB_TOP - DB_BOTTOM;
        if (len <= 0) return;
        // 割合を座標にする (右から左は反転)
        java.util.function.DoubleUnaryOperator px = f -> b.rtl() ? left + (1 - f) * len : left + f * len;
        double half = DB_AXIS_W / 2;
        double a = px.applyAsDouble(b.start()), e = px.applyAsDouble(b.end());
        double x0 = Math.min(a, e), x1 = Math.max(a, e);
        if (!Double.isNaN(b.axis())) {
            // 軸の線の外側から棒を伸ばす
            double ax = px.applyAsDouble(b.axis());
            if (x0 >= ax - 0.01) x0 = Math.max(x0, ax + half);
            if (x1 <= ax + 0.01) x1 = Math.min(x1, ax - half);
        }
        if (x1 - x0 > 0.01 && bh > 0) {
            if (b.gradient()) {
                // 濃い側: 軸・始まりの側。反対側はバーの色を 85% 白に寄せた色
                boolean darkLeft = b.negative() == b.rtl();
                Color light = lighten(b.fill(), 0.85);
                cv.fillAxial(x0, top, x1 - x0, bh, darkLeft ? x0 : x1, top, darkLeft ? x1 : x0, top, b.fill(), light);
            } else {
                cv.fillRect(x0, top, x1 - x0, bh, b.fill());
            }
            if (b.border() != null) {
                double t = DB_BORDER;
                cv.fillRect(x0, top, x1 - x0, t, b.border());
                cv.fillRect(x0, top + bh - t, x1 - x0, t, b.border());
                cv.fillRect(x0, top + t, t, bh - 2 * t, b.border());
                cv.fillRect(x1 - t, top + t, t, bh - 2 * t, b.border());
            }
        }
        if (!Double.isNaN(b.axis())) {
            // 軸: セルの高さいっぱいの破線 (2.88pt 描いて 0.96pt 空ける)
            double ax = px.applyAsDouble(b.axis());
            for (double yy = y; yy < y + h - 0.01; yy += DB_DASH + DB_GAP) {
                cv.fillRect(ax - half, yy, DB_AXIS_W, Math.min(DB_DASH, y + h - yy), b.axisColor());
            }
        }
    }

    /** データバーを描ける範囲の余白・枠線・軸の寸法 (pt)。 */
    static final double DB_START = 1.92, DB_END = 0.96, DB_TOP = 1.92, DB_BOTTOM = 2.24, DB_BORDER = 0.96;
    static final double DB_AXIS_W = 0.96, DB_DASH = 2.88, DB_GAP = 0.96;

    /** 色を白に寄せる (f = 0 で元の色、1 で白)。 */
    private static Color lighten(Color c, double f) {
        return new Color((int) Math.round(c.getRed() + (255 - c.getRed()) * f), (int) Math.round(c.getGreen() + (255 - c.getGreen()) * f),
                (int) Math.round(c.getBlue() + (255 - c.getBlue()) * f));
    }

    // -------------------------------------------------------------- 行列番号

    /**
     * 行列番号 (printOptions headings) を描く。座標はシートの pt で、ブロックの左上が (L.x(s.c0), L.y(s.r0))。
     * 白地に黒の枠 (用紙上 0.96pt)。列名は中央、行番号は横中央・縦は下揃え。
     *
     * @param top  ページの一番上のブロック (列の見出しを描く)
     * @param left ページの一番左のブロック (行番号を描く)
     */
    public void paintHeadings(PdfCanvas cv, Paginator.Segment s, double hdrW, double hdrH, boolean top, boolean left,
                              boolean right, boolean bottom) throws IOException {
        double lw = 0.96 / Math.max(0.01, L.scale);
        Color line = Color.BLACK;
        double sx = L.x(s.c0), sy = L.y(s.r0);
        FontManager.Spec spec = headingFont();
        float size = headingSize();
        // 游ゴシック (9pt 超) の行列番号は Meiryo UI 9pt で描き、Excel はそれを上・右へずらす
        String dn = ctx.defaultFontName == null ? "" : ctx.defaultFontName;
        boolean yuHead = (dn.startsWith("游ゴシック") || dn.toLowerCase(java.util.Locale.ROOT).startsWith("yu gothic")) && ctx.defaultFontSize > 9;
        double asc = spec.ascent * size, desc = spec.descent * size;
        if (top) {
            for (int c = s.c0; c <= s.c1; c++) {
                double w = L.cw(c);
                if (w <= 0) continue;
                double x = L.x(c);
                cv.fillRect(x, sy - hdrH, w, hdrH, Color.WHITE);
                String lab = org.apache.poi.ss.util.CellReference.convertNumToColString(c);
                double tw = spec.width(lab, size, ctx.fonts);
                cv.text(lab, x + (w - tw) / 2 + (yuHead ? YU_HEAD_DX : 0), sy - hdrH + (hdrH - asc - desc) / 2 + asc - (yuHead ? YU_HEAD_DY : 0), spec, size, Color.BLACK, 1, 0);
                cv.fillRect(x - lw / 2, sy - hdrH, lw, hdrH, line);
                cv.fillRect(x + w - lw / 2, sy - hdrH, lw, hdrH, line);
            }
            double x1 = L.x(s.c1 + 1);
            cv.fillRect(sx - (left ? hdrW : 0), sy - hdrH - lw / 2, x1 - sx + (left ? hdrW : 0), lw, line);
            cv.fillRect(sx - (left ? hdrW : 0), sy - lw / 2, x1 - sx + (left ? hdrW : 0), lw, line);
        }
        if (left) {
            for (int r = s.r0; r <= s.r1; r++) {
                double h = L.rh(r);
                if (h <= 0) continue;
                double y = L.y(r);
                cv.fillRect(sx - hdrW, y, hdrW, h, Color.WHITE);
                String lab = String.valueOf(r + 1);
                double tw = spec.width(lab, size, ctx.fonts);
                // 基線はセルの文字 (標準フォントの下揃え) と同じ高さ
                FontManager.Spec ds = ctx.fonts.spec(ctx.defaultFontName, false, false);
                double base = y + h - padBottom() - ds.descent * ctx.defaultFontSize;
                cv.text(lab, sx - hdrW + (hdrW - tw) / 2 + (yuHead ? YU_ROWNUM_DX : 0), base - (yuHead ? YU_ROWNUM_DY : 0), spec, size, Color.BLACK, 1, 0);
                cv.fillRect(sx - hdrW, y - lw / 2, hdrW, lw, line);
                cv.fillRect(sx - hdrW, y + h - lw / 2, hdrW, lw, line);
            }
            double y1 = L.y(s.r1 + 1);
            cv.fillRect(sx - hdrW - lw / 2, sy - (top ? hdrH : 0), lw, y1 - sy + (top ? hdrH : 0), line);
            cv.fillRect(sx - lw / 2, sy - (top ? hdrH : 0), lw, y1 - sy + (top ? hdrH : 0), line);
        }
        // 印刷する範囲の右端と下端の枠
        double x1 = L.x(s.c1 + 1), y1 = L.y(s.r1 + 1);
        if (right) cv.fillRect(x1 - lw / 2, sy - (top ? hdrH : 0), lw, y1 - sy + (top ? hdrH : 0), line);
        if (bottom) cv.fillRect(sx - (left ? hdrW : 0), y1 - lw / 2, x1 - sx + (left ? hdrW : 0), lw, line);
    }

    /** 行列番号のフォント: 標準フォント。ただし游ゴシック (9pt を超える) は Meiryo UI (Excel の印刷)。 */
    static final double YU_HEAD_DX = 0.93, YU_HEAD_DY = 2.66, YU_ROWNUM_DX = 0.69, YU_ROWNUM_DY = 0.2;

    private FontManager.Spec headingFont() {
        String n = ctx.defaultFontName == null ? "" : ctx.defaultFontName;
        boolean yu = n.startsWith("游ゴシック") || n.toLowerCase(java.util.Locale.ROOT).startsWith("yu gothic");
        if (yu && ctx.defaultFontSize > 9) return ctx.fonts.spec("Meiryo UI", false, false);
        return ctx.fonts.spec(n, false, false);
    }

    /** 行列番号の文字の大きさ (pt)。 */
    private float headingSize() {
        String n = ctx.defaultFontName == null ? "" : ctx.defaultFontName;
        boolean yu = n.startsWith("游ゴシック") || n.toLowerCase(java.util.Locale.ROOT).startsWith("yu gothic");
        return (float) (yu && ctx.defaultFontSize > 9 ? 9 : ctx.defaultFontSize);
    }

    // -------------------------------------------------------------- 枠線

    private void paintGridlines(PdfCanvas cv, Paginator.Segment s) throws IOException {
        // Excel の印刷は枠線を 0.14pt の線と、同じ位置の 0.12pt の塗りの重ねで描く (60dpi で B5 の濃さ)
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.14f);
        Color gc = Color.BLACK;
        for (int r = s.r0; r <= s.r1 + 1; r++) {
            double y = L.y(r);
            int c = s.c0;
            while (c <= s.c1) {
                int start = c;
                while (c <= s.c1 && !insideMergedH(r, c)) c++;
                if (c > start) {
                    cv.line(L.x(start), y, L.x(c), y, st, gc);
                    cv.fillRect(L.x(start), y - 0.06, L.x(c) - L.x(start), 0.12, gc);
                }
                while (c <= s.c1 && insideMergedH(r, c)) c++;
            }
        }
        for (int cc = s.c0; cc <= s.c1 + 1; cc++) {
            double x = L.x(cc);
            int r = s.r0;
            while (r <= s.r1) {
                int start = r;
                while (r <= s.r1 && !insideMergedV(r, cc)) r++;
                if (r > start) {
                    cv.line(x, L.y(start), x, L.y(r), st, gc);
                    cv.fillRect(x - 0.06, L.y(start), 0.12, L.y(r) - L.y(start), gc);
                }
                while (r <= s.r1 && insideMergedV(r, cc)) r++;
            }
        }
    }

    /** 行 r の上・列 c の位置の横の枠線が、結合範囲の内側か。 */
    private boolean insideMergedH(int r, int c) {
        CellRangeAddress m = L.mergedAt(r, c);
        return m != null && r > m.getFirstRow();
    }

    /** 列 c の左・行 r の位置の縦の枠線が、結合範囲の内側か。 */
    private boolean insideMergedV(int r, int c) {
        CellRangeAddress m = L.mergedAt(r, c);
        return m != null && c > m.getFirstColumn();
    }

    // -------------------------------------------------------------- 罫線

    /** 隣り合うセルの罫線が重なったときの優先度 (大きいほうを描く)。 */
    static int weight(BorderStyle s) {
        return switch (s) {
            case NONE -> 0;
            case HAIR -> 1;
            case DOTTED -> 2;
            case DASH_DOT_DOT -> 3;
            case DASH_DOT -> 4;
            case DASHED -> 5;
            case THIN -> 6;
            case MEDIUM_DASH_DOT_DOT -> 7;
            case SLANTED_DASH_DOT -> 8;
            case MEDIUM_DASH_DOT -> 9;
            case MEDIUM_DASHED -> 10;
            case MEDIUM -> 11;
            case THICK -> 12;
            case DOUBLE -> 13;
        };
    }

    /** セルの罫線の 1 辺 (0 上、1 右、2 下、3 左。条件付き書式を優先。なければ null)。 */
    private Edge edge(XSSFCellStyle st, int side, CondFormat.Result cfr) {
        BorderStyle bs;
        XSSFColor xc;
        if (cfr != null && cfr.borders != null && cfr.borders[side] != null) return cfr.borders[side];
        if (st == null) return null;
        switch (side) {
            case 0 -> { bs = st.getBorderTop(); xc = st.getTopBorderXSSFColor(); }
            case 1 -> { bs = st.getBorderRight(); xc = st.getRightBorderXSSFColor(); }
            case 2 -> { bs = st.getBorderBottom(); xc = st.getBottomBorderXSSFColor(); }
            default -> { bs = st.getBorderLeft(); xc = st.getLeftBorderXSSFColor(); }
        }
        if (bs == null || bs == BorderStyle.NONE) return null;
        Color c = ctx.theme.color(xc);
        return new Edge(bs, c == null || bw() ? Color.BLACK : c);
    }

    /** セルの罫線の 1 辺。セルの書式になければテーブルのスタイルの罫線。 */
    private Edge edge(XSSFCellStyle st, int side, CondFormat.Result cfr, int r, int c) {
        Edge e = edge(st, side, cfr);
        return e != null ? e : tables.edge(r, c, side, bw());
    }

    /**
     * 優先度の高いほうの罫線。同じ太さなら暗い色、同じ明るさなら後のセル (下・右) の罫線。
     */
    private static Edge better(Edge a, Edge b) {
        if (a == null) return b;
        if (b == null) return a;
        // 条件付き書式で消した辺は、隣のセルの罫線より優先する
        if (a.cleared()) return a;
        if (b.cleared()) return b;
        // セル (条件付き書式を含む) の罫線は、隣のセルの表のスタイルの罫線より優先する (太さによらない。Excel の実測)
        if (a.table() != b.table()) return a.table() ? b : a;
        int wa = weight(a.style()), wb = weight(b.style());
        if (wb == wa) {
            // 同じ太さなら暗い色 (輝度の低い色。位置によらない。黒 > 灰色、青 > 赤 > 薄い灰色)。同じ明るさなら後のセル
            if (a.color() != null && b.color() != null) {
                int la = luma(a.color()), lb = luma(b.color());
                if (la != lb) return la < lb ? a : b;
            }
            return b;
        }
        return wb > wa ? b : a;
    }


    private static int luma(Color c) { return 299 * c.getRed() + 587 * c.getGreen() + 114 * c.getBlue(); }

    /** 罫線を描く。隣り合うセルの辺をまとめ、同じ罫線が続く区間を 1 本の線にする。 */
    private void paintBorders(PdfCanvas cv, Paginator.Segment s) throws IOException {
        int nr = s.r1 - s.r0 + 1, nc = s.c1 - s.c0 + 1;
        Edge[][] hz = new Edge[nr + 1][nc]; // hz[i][j]: 行 s.r0+i の上・列 j の横線
        Edge[][] vt = new Edge[nr][nc + 1]; // vt[i][j]: 行 i・列 s.c0+j の左の縦線
        List<double[]> diagonals = new ArrayList<>();
        List<Edge> diagEdges = new ArrayList<>();
        for (int r = s.r0; r <= s.r1; r++) {
            ctx.checkDeadline();
            if (L.rh(r) <= 0) continue;
            for (int c = s.c0; c <= s.c1; c++) {
                if (L.cw(c) <= 0) continue;
                XSSFCellStyle st = style(r, c);
                CondFormat.Result cfr = cf.eval(r, c);
                if (st == null && (cfr == null || cfr.borders == null) && !tables.any()) continue;
                CellRangeAddress m = L.mergedAt(r, c);
                int i = r - s.r0, j = c - s.c0;
                boolean top = m == null || r == m.getFirstRow();
                boolean bottom = m == null || r == m.getLastRow();
                boolean left = m == null || c == m.getFirstColumn();
                boolean right = m == null || c == m.getLastColumn();
                double slant = st == null || m != null ? 0 : slantAngle(st);
                if (slant != 0) {
                    slantEdges(st, cfr, r, c, slant, diagonals, diagEdges);
                    continue;
                }
                if (top) hz[i][j] = better(hz[i][j], edge(st, 0, cfr, r, c));
                if (bottom) hz[i + 1][j] = better(hz[i + 1][j], edge(st, 2, cfr, r, c));
                if (left) vt[i][j] = better(vt[i][j], edge(st, 3, cfr, r, c));
                if (right) vt[i][j + 1] = better(vt[i][j + 1], edge(st, 1, cfr, r, c));
                if (st != null && (m == null || (r == m.getFirstRow() && c == m.getFirstColumn()))) {
                    collectDiagonal(st, m, r, c, diagonals, diagEdges);
                }
            }
        }
        // 前のページの列にある右へ傾けたセルの平行四辺形の続き
        if (s.c0 > 0) for (int r = s.r0; r <= s.r1; r++) {
            var row = sheet.getRow(r);
            if (row == null || L.rh(r) <= 0) continue;
            for (var cl : row) {
                int c = cl.getColumnIndex();
                if (c >= s.c0 || !(cl.getCellStyle() instanceof XSSFCellStyle cst) || L.mergedAt(r, c) != null) continue;
                double sl = slantAngle(cst);
                if (sl > 0 && L.x(c) + L.cw(c) + L.rh(r) / Math.tan(Math.toRadians(sl)) > L.x(s.c0))
                    slantEdges(cst, cf.eval(r, c), r, c, sl, diagonals, diagEdges);
            }
        }
        // 次のページの列にある左へ傾けたセルの平行四辺形
        for (int r = s.r0; r <= s.r1; r++) {
            var row = sheet.getRow(r);
            if (row == null || L.rh(r) <= 0) continue;
            for (var cl : row) {
                int c = cl.getColumnIndex();
                if (c <= s.c1 || !(cl.getCellStyle() instanceof XSSFCellStyle cst) || L.mergedAt(r, c) != null) continue;
                double sl = slantAngle(cst);
                if (sl < 0 && L.x(c) - L.rh(r) / Math.tan(Math.toRadians(-sl)) < L.x(s.c1 + 1))
                    slantEdges(cst, cf.eval(r, c), r, c, sl, diagonals, diagEdges);
            }
        }
        // 印刷範囲がなければ、ブロックの外周の罫線は隣のセルのものでもよい。印刷範囲があると、Excel は範囲の外のセルの
        // 罫線を描かない。印刷範囲内の改ページの向こう側のセルの罫線も印刷しない
        CellRangeAddress a = s.area;
        for (int c = s.c0; c <= s.c1; c++) {
            if (L.cw(c) <= 0) continue;
            int j = c - s.c0;
            if (s.r0 > 0 && a == null) hz[0][j] = better(neighbour(s.r0 - 1, c, 2), hz[0][j]);   // 同じ太さなら下 (ブロックの中) のセル
            if (a == null) hz[nr][j] = better(hz[nr][j], neighbour(s.r1 + 1, c, 0));
        }
        for (int r = s.r0; r <= s.r1; r++) {
            ctx.checkDeadline();
            if (L.rh(r) <= 0) continue;
            int i = r - s.r0;
            if (s.c0 > 0 && a == null) vt[i][0] = better(neighbour(r, s.c0 - 1, 1), vt[i][0]);   // 同じ太さなら右 (ブロックの中) のセル
            if (a == null) vt[i][nc] = better(vt[i][nc], neighbour(r, s.c1 + 1, 3));
        }
        // 「選択範囲内で中央」でつながった範囲 (文字のあるセルと、それに続く同じ配置の空のセル) は、結合セルと同様に
        // 内側の縦罫線を描かない
        for (int r = s.r0; r <= s.r1; r++) {
            int i = r - s.r0;
            for (int c = s.c0; c <= s.c1; c++) {
                int end = centerSelectionEnd(r, c);
                for (int cc = c + 1; cc <= end; cc++) if (cc - s.c0 <= nc && cc > s.c0) vt[i][cc - s.c0] = null;
                if (end > c) c = end;
            }
        }
        for (int i = 0; i < nr && !overflowEdges.isEmpty(); i++)
            for (int j = 0; j <= nc; j++) if (overflowEdges.contains((s.r0 + i) * 100000L + s.c0 + j)) vt[i][j] = null;
        // 条件付き書式で消した辺は描かない
        for (Edge[] row : hz) for (int j = 0; j < row.length; j++) if (row[j] != null && row[j].cleared()) row[j] = null;
        for (Edge[] row : vt) for (int j = 0; j < row.length; j++) if (row[j] != null && row[j].cleared()) row[j] = null;
        // 横線 (同じ罫線が続く区間ごと)
        for (int i = 0; i <= nr; i++) {
            int r = s.r0 + i;
            double y = L.y(r);
            int j = 0;
            while (j < nc) {
                Edge e = hz[i][j];
                if (e == null) { j++; continue; }
                int k = j;
                while (k + 1 < nc && e.equals(hz[i][k + 1])) k++;
                drawEdge(cv, e, L.x(s.c0 + j), y, L.x(s.c0 + k + 1), y, true,
                        vt, i, j, k, nr);
                j = k + 1;
            }
        }
        // 縦線
        for (int j = 0; j <= nc; j++) {
            int c = s.c0 + j;
            double x = L.x(c);
            int i = 0;
            while (i < nr) {
                Edge e = vt[i][j];
                if (e == null) { i++; continue; }
                int k = i;
                while (k + 1 < nr && e.equals(vt[k + 1][j])) k++;
                drawEdge(cv, e, x, L.y(s.r0 + i), x, L.y(s.r0 + k + 1), false, null, 0, 0, 0, 0);
                i = k + 1;
            }
        }
        // 斜めの線 (対角線と、斜めの文字のセルの平行四辺形) は、ページの本体の左右の範囲で切る
        boolean clipDiag = !diagonals.isEmpty() && cv != null;
        if (clipDiag) {
            cv.save();
            cv.clipRect(L.x(s.c0), L.y(s.r0) - 10000, L.x(s.c1 + 1) - L.x(s.c0), 20000 + L.y(s.r1 + 1) - L.y(s.r0));
        }
        try {
            for (int d = 0; d < diagonals.size(); d++) {
                double[] p = diagonals.get(d);
                drawEdge(cv, diagEdges.get(d), p[0], p[1], p[2], p[3], false, null, 0, 0, 0, 0);
            }
        } finally {
            if (clipDiag) cv.restore();
        }
    }

    /**
     * セル (r, c) が文字のある「選択範囲内で中央」のセルなら、続く同じ配置の空のセルの最後の列を返す。そうでなければ c。
     * 結合セルにかかる場合は広げない。
     */
    private int centerSelectionEnd(int r, int c) {
        XSSFCell cell = cell(r, c);
        if (cell == null || isEmpty(cell) || L.mergedAt(r, c) != null) return c;
        if (align(cell.getCellStyle()).h() != HorizontalAlignment.CENTER_SELECTION) return c;
        int cc = c + 1;
        while (cc < L.maxCol && isEmpty(cell(r, cc)) && L.mergedAt(r, cc) == null) {
            XSSFCellStyle ns = style(r, cc);
            if (ns == null || align(ns).h() != HorizontalAlignment.CENTER_SELECTION) break;
            cc++;
        }
        return cc - 1;
    }

    /** ブロックの外のセルの、ブロックに面した辺の罫線 (0 上、1 右、2 下、3 左)。 */
    private Edge neighbour(int r, int c, int side) {
        CellRangeAddress m = L.mergedAt(r, c);
        if (m != null) {
            boolean onEdge = switch (side) {
                case 0 -> r == m.getFirstRow();
                case 2 -> r == m.getLastRow();
                case 3 -> c == m.getFirstColumn();
                default -> c == m.getLastColumn();
            };
            if (!onEdge) return null;
        }
        XSSFCellStyle st = style(r, c);
        if (st != null && m == null && slantAngle(st) != 0) return null; // 斜めの文字のセルの罫線は平行四辺形で描く
        CondFormat.Result cfr = cf.eval(r, c);
        if (st == null && (cfr == null || cfr.borders == null) && !tables.any()) return null;
        return edge(st, side, cfr, r, c);
    }

    /** セル (結合範囲) の斜線を集める。 */
    private void collectDiagonal(XSSFCellStyle st, CellRangeAddress m, int r, int c, List<double[]> out, List<Edge> edges) {
        CTXf xf = st.getCoreXf();
        if (xf == null || !xf.isSetBorderId()) return;
        CTBorder b = ctx.wb.getStylesSource().getBorderAt((int) xf.getBorderId()).getCTBorder();
        if (b == null || !b.isSetDiagonal()) return;
        CTBorderPr d = b.getDiagonal();
        if (!d.isSetStyle()) return;
        BorderStyle bs = BorderStyle.valueOf((short) (d.getStyle().intValue() - 1));
        if (bs == BorderStyle.NONE) return;
        Color col = Color.BLACK;
        if (d.isSetColor()) {
            XSSFColor xc = XSSFColor.from(d.getColor(), ctx.wb.getStylesSource().getIndexedColors());
            if (xc != null && xc.isThemed()) ctx.wb.getStylesSource().getTheme().inheritFromThemeAsRequired(xc);
            Color cc = ctx.theme.color(xc);
            if (cc != null && !bw()) col = cc;
        }
        double x0 = L.x(c), y0 = L.y(r), x1, y1;
        if (m != null) { x1 = L.x(m.getLastColumn() + 1); y1 = L.y(m.getLastRow() + 1); }
        else { x1 = x0 + L.cw(c); y1 = y0 + L.rh(r); }
        Edge e = new Edge(bs, col);
        if (b.isSetDiagonalDown() && b.getDiagonalDown()) { out.add(new double[] {x0, y0, x1, y1}); edges.add(e); }
        if (b.isSetDiagonalUp() && b.getDiagonalUp()) { out.add(new double[] {x0, y1, x1, y0}); edges.add(e); }
    }

    /** 文字を斜めに回したセルの罫線: 文字の向きに傾けた平行四辺形の 4 辺。 */
    private void slantEdges(XSSFCellStyle st, CondFormat.Result cfr, int r, int c, double slant, List<double[]> out, List<Edge> edges) {
        double x0 = L.x(c), x1 = x0 + L.cw(c), yT = L.y(r), yB = yT + L.rh(r);
        double sh = L.rh(r) / Math.tan(Math.toRadians(Math.abs(slant))) * Math.signum(slant);
        double[][] seg = {{x0 + sh, yT, x1 + sh, yT}, {x1, yB, x1 + sh, yT}, {x0, yB, x1, yB}, {x0, yB, x0 + sh, yT}};
        for (int k = 0; k < 4; k++) {
            Edge e = edge(st, k, cfr, r, c);
            if (e != null) { out.add(seg[k]); edges.add(e); }
        }
        collectDiagonal(st, null, r, c, out, edges);
    }

    /** ブロックの左の列にある右へ傾けた文字のセルの平行四辺形が、ブロックにかかるか (そのページは空でない)。 */
    public boolean hasSlantSpill(Paginator.Segment b) {
        for (int r = b.r0; r <= b.r1; r++) {
            var row = sheet.getRow(r);
            if (row == null || L.rh(r) <= 0) continue;
            for (var cl : row) {
                int c = cl.getColumnIndex();
                if (c >= b.c0 || !(cl.getCellStyle() instanceof XSSFCellStyle cst)) continue;
                double sl = slantAngle(cst);
                if (sl <= 0 || L.mergedAt(r, c) != null || !hasAnyBorder(cst)) continue;
                if (L.x(c) + L.cw(c) + L.rh(r) / Math.tan(Math.toRadians(sl)) > L.x(b.c0)) return true;
            }
        }
        return false;
    }

    /** 収まらない数値の # の個数を決めるとき、余白の内側にはみ出してよい幅 (600 dpi の画素)。 */
    static final int HASH_OVER_PX = 2;

    /**
     * 収まらない数値の # の並び (n 文字)。分数 (? の空白 U+2007 と / を含む表示) は、表示の文字列の分子の前の空白だけを空白のまま残し、
     * ほかの文字 (整数部とその後の区切りの空白を含む) を # にして先頭から n 文字 (「# ??/??」の 12.4 は「### ##」)。
     */
    static String fractionHashes(String text, int n) {
        int slash = text == null ? -1 : text.indexOf('/');
        if (slash < 0 || text.indexOf('\u2007') < 0) return "#".repeat(n);
        int k = slash - 1;
        while (k >= 0 && Character.isDigit(text.charAt(k))) k--;
        int padEnd = k; // 分子の前の空白の終わり
        while (k >= 0 && text.charAt(k) == '\u2007') k--;
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < text.length() && sb.length() < n; i++) sb.append(i > k && i <= padEnd ? '\u2007' : '#');
        while (sb.length() < n) sb.append('#');
        return sb.toString();
    }

    /** 罫線 (上下左右) のある書式か (斜めの文字のセルは、罫線があるときだけ平行四辺形の分を印刷の範囲に入れる)。 */
    private static boolean hasAnyBorder(XSSFCellStyle st) {
        return st.getBorderTop() != BorderStyle.NONE || st.getBorderBottom() != BorderStyle.NONE
                || st.getBorderLeft() != BorderStyle.NONE || st.getBorderRight() != BorderStyle.NONE;
    }

    /** 文字の角度が斜め (0 と ±90 以外) なら、その角度 (度。正は反時計回り)。それ以外は 0。 */
    private double slantAngle(XSSFCellStyle st) {
        int rot = align(st).rotation();
        if (rot == 0 || rot == 255 || rot == 90 || rot == 180) return 0;
        return rot < 90 ? rot : -(rot - 90);
    }

    /** 罫線の太さ (pt)。 */
    static float lineWidth(BorderStyle s) {
        return switch (s) {
            case HAIR -> 0.24f;
            case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT -> 1.8f;
            case THICK -> 2.8f;
            default -> THIN;
        };
    }

    /** 罫線の破線のパターン (実線なら null)。 */
    static float[] dash(BorderStyle s) {
        float u = THIN;
        return switch (s) {
            case DOTTED -> new float[] {u, u};
            case DASHED -> new float[] {3 * u, u};
            case DASH_DOT -> new float[] {6 * u, 2 * u, 2 * u, 2 * u};
            case DASH_DOT_DOT -> new float[] {6 * u, 2 * u, 2 * u, 2 * u, 2 * u, 2 * u};
            case MEDIUM_DASHED -> new float[] {6 * u, 3 * u};
            case MEDIUM_DASH_DOT, SLANTED_DASH_DOT -> new float[] {6 * u, 2 * u, 2 * u, 2 * u};
            case MEDIUM_DASH_DOT_DOT -> new float[] {6 * u, 2 * u, 2 * u, 2 * u, 2 * u, 2 * u};
            default -> null;
        };
    }

    /** 罫線を 1 本描く (二重線は 2 本の細線)。 */
    private void drawEdge(PdfCanvas cv, Edge e, double x0, double y0, double x1, double y1, boolean horizontal,
                          Edge[][] vt, int i, int j0, int j1, int nr) throws IOException {
        BorderStyle bs = e.style();
        if (posx() && !L.model().authorEnv && bs != BorderStyle.SLANTED_DASH_DOT
                && (x0 == x1 || y0 == y1)) {
            BorderLines.draw(cv, e, x0, y0, x1, y1, L.scale, blockOx - blockSx * L.scale, blockOy - blockSy * L.scale);
            return;
        }
        float w = lineWidth(bs);
        float k = (float) (1 / Math.max(0.1, L.scale)); // 用紙上の pt → シートの単位
        if (bs == BorderStyle.HAIR) w = 0.12f * k;      // Excel: プリンタの 1 ピクセルの実線
        if (bs == BorderStyle.DOUBLE) {
            float lw = THIN * 0.8f;
            double off = THIN;
            PdfCanvas.Stroke st = new PdfCanvas.Stroke(lw);
            st.cap = 2;
            if (horizontal || y0 == y1) {
                cv.line(x0, y0 - off, x1, y1 - off, st, e.color());
                cv.line(x0, y0 + off, x1, y1 + off, st, e.color());
            } else if (x0 == x1) {
                cv.line(x0 - off, y0, x1 - off, y1, st, e.color());
                cv.line(x0 + off, y0, x1 + off, y1, st, e.color());
            } else {
                cv.line(x0, y0, x1, y1, st, e.color());
            }
            return;
        }
        float[] d = dash(bs);
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(w);
        // Excel は点線を、半分がインクの 8x8 プリンタピクセルのタイルで塗る: 0.48pt 描いて 0.48pt 空ける
        if (bs == BorderStyle.DOTTED) d = new float[] {0.48f * k, 0.48f * k};
        if (d != null) st.dash(d);
        else st.cap = 2; // 四角い線端で角を閉じる
        if (posx() && d == null && bs != BorderStyle.HAIR && (x0 == x1 || y0 == y1)) {
            // Excel は実線を塗りつぶしの長方形で出力する (re f*)。太さは 8 / 16 / 24 画素 × 倍率を丸めた値。線 (stroke) は表示側で画素に寄せられ、1 列ずれることがある
            int base = bs == BorderStyle.THIN ? 8 : bs == BorderStyle.THICK ? 24 : 16;
            float wf = (float) (Math.floor(base * L.scale + 0.5) * 0.12 / Math.max(0.1, L.scale));
            double hw = wf / 2;
            java.awt.Color col = e.color() != null ? e.color() : java.awt.Color.BLACK;   // 自動の色は黒
            if (y0 == y1) cv.fillRect(Math.min(x0, x1) - hw, y0 - hw, Math.abs(x1 - x0) + wf, wf, col);
            else cv.fillRect(x0 - hw, Math.min(y0, y1) - hw, wf, Math.abs(y1 - y0) + wf, col);
        } else cv.line(x0, y0, x1, y1, st, e.color());
        // Excel は倍率 100% の細い罫線を、0.14pt の線と 0.96pt の帯の 2 本で描く
        if (L.scale == 1 && bs == BorderStyle.THIN && !L.model().authorEnv) {
            PdfCanvas.Stroke h = new PdfCanvas.Stroke(0.14f);
            if (horizontal || y0 == y1) cv.line(x0, y0 - 0.42, x1, y1 - 0.42, h, e.color());
            else if (x0 == x1) cv.line(x0 - 0.42, y0, x1 - 0.42, y1, h, e.color());
        }
    }

    // -------------------------------------------------------------- 文字

    /**
     * Excel がプリンタで実体化するフォントサイズ: サイズ × 倍率を整数デバイスピクセルに丸めたもの
     * (10pt を 75% / 600 dpi で 62.5 → 63 px、シートの単位で 10.08pt)。折り返し位置はこれで決まる。
     */
    float realize(double sizePt) {
        // 縦書きの文字は実体化しない
        double k = (!stackedText && !L.model().authorEnv ? L.scale : L.realScale()) * ps.vdpi / 72.0;
        if (k <= 0) return (float) sizePt;
        double px = Math.round(sizePt * k);
        // Excel は大きさを 1/20pt で丸めてから画素にする: round(round(pt × 20 × 倍率) × 5 / 12) (11pt × 77% は 70 画素、11pt × 68% は 63 画素)
        if (!stackedText && !L.model().authorEnv && ps.vdpi == 600) {
            px = Math.floor(Math.floor(sizePt * 20 * L.scale + 0.5 + 1e-9) * 5 / 12 + 0.5 + 1e-9);
        }
        return (float) (px <= 0 ? sizePt : px / k);
    }

    /**
     * セルのフォントに対するフォント。リッチテキストのランのフォントは styles.xml に登録されておらず、
     * どれも同じ番号 (getIndex) を返すため、番号ではなく名前・太字・斜体で選ぶ (FontManager がキャッシュする)。
     * 番号でキャッシュすると、ランが最初にキャッシュされた別のフォントで描かれる。
     */
    FontManager.Spec spec(XSSFFont f) {
        return ctx.fonts.spec(ctx.cellFontName(f), f.getBold(), f.getItalic());
    }

    /** 白黒印刷: 塗りつぶしなし、文字と罫線は黒。 */
    private boolean bw() { return ps.blackAndWhite; }

    /** フォントの色 (自動なら黒)。 */
    Color fontColor(XSSFFont f) {
        Color c = ctx.theme.color(f.getXSSFColor());
        return c == null ? Color.BLACK : c;
    }

    /** セルのフォントの色がテーブルのスタイルの文字色より優先するか。 */
    private boolean explicitFontColor(XSSFFont f) {
        Color c = ctx.theme.color(f.getXSSFColor());
        return c != null && !c.equals(ctx.theme.color(ctx.wb.getFontAt(0).getXSSFColor()));
    }

    /** 書式 (フォント、条件付き書式、表示形式の色) を適用したランを作る。 */
    private TextLayout.Run run(String text, XSSFFont f, CondFormat.Result cfr, Color fmtColor) {
        boolean bold = f.getBold(), italic = f.getItalic();
        TableStyles.Look lk = runLook;
        // テーブルの要素の太字・斜体は、セルのフォントになければ足す (要素で消した val="0" はセルのフォントを消さない)
        if (lk != null && Boolean.TRUE.equals(lk.bold)) bold = true;
        if (lk != null && Boolean.TRUE.equals(lk.italic)) italic = true;
        if (cfr != null && cfr.bold != null) bold = cfr.bold;
        if (cfr != null && cfr.italic != null) italic = cfr.italic;
        FontManager.Spec sp = (bold == f.getBold() && italic == f.getItalic()) ? spec(f)
                : ctx.fonts.spec(ctx.cellFontName(f), bold, italic);
        Color col = fmtColor != null ? fmtColor : (cfr != null && cfr.fontColor != null ? cfr.fontColor
                : lk != null && lk.font != null && !explicitFontColor(f) ? lk.font : fontColor(f));
        if (bw()) col = Color.BLACK;
        TextLayout.Run r = new TextLayout.Run(text, sp, realize(f.getFontHeight() / 20.0), col);
        byte u = f.getUnderline();
        r.underline = u == 0 ? 0 : (u == 2 || u == 34) ? 2 : 1;
        r.accounting = u == 33 || u == 34;
        if (u == 0 && lk != null && lk.underline != null) r.underline = lk.underline;
        r.strike = f.getStrikeout() || (cfr != null && Boolean.TRUE.equals(cfr.strike)) || (lk != null && Boolean.TRUE.equals(lk.strike));
        r.baselineShift = f.getTypeOffset() == 1 ? 1 : f.getTypeOffset() == 2 ? -1 : 0;
        r.cellShift = true;
        return r;
    }

    /** セルの文字列のラン (書式付き文字列なら区切りごと)。 */
    private List<TextLayout.Run> runs(ValueFormatter.Result vr, XSSFFont base, CondFormat.Result cfr) {
        List<TextLayout.Run> out = new ArrayList<>();
        if (vr.rich != null && vr.fillChar == null) {
            XSSFRichTextString rs = vr.rich;
            String s = rs.getString();
            int n = rs.numFormattingRuns();
            int pos = 0;
            for (int i = 0; i < n; i++) {
                // ランの位置は文字列の長さに収める (_x0000_ などの制御文字の数え方で POI の文字列より先を指すことがある)
                int start = Math.min(s.length(), rs.getIndexOfFormattingRun(i));
                int len = rs.getLengthOfFormattingRun(i);
                if (start > pos) out.add(run(s.substring(pos, start), base, cfr, vr.color));
                XSSFFont f = rs.getFontOfFormattingRun(i);
                if (f == null) f = base;
                int end = Math.min(s.length(), start + len);
                if (end > start) out.add(run(s.substring(start, end), f, cfr, vr.color));
                pos = Math.max(pos, end);
            }
            if (pos < s.length()) out.add(run(s.substring(pos), base, cfr, vr.color));
            return out;
        }
        out.add(run(vr.text, base, cfr, vr.color));
        return out;
    }

    /** ブロックのセルの文字を描く。このブロックの左右のセルからはみ出す文字は、Excel は印刷しない。 */
    private void paintTexts(PdfCanvas cv, Paginator.Segment s) throws IOException {
        if (s.textDx != 0 || s.textDy != 0) {
            cv.save();
            cv.translate(s.textDx, s.textDy);
            try {
                paintTexts0(cv, s);
            } finally {
                cv.restore();
            }
        } else paintTexts0(cv, s);
    }

    private void paintTexts0(PdfCanvas cv, Paginator.Segment s) throws IOException {
        int depth = cv == null ? 0 : cv.depth();   // セルの例外で途中のクリップ・変換が残らないよう、ここまで戻す
        for (int r = s.r0; r <= s.r1; r++) {
            ctx.checkDeadline();
            if (L.rh(r) <= 0) continue;
            XSSFRow row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = s.c0; c <= s.c1; c++) {
                CellRangeAddress m = L.mergedAt(r, c);
                if (m != null) {
                    if (r != firstVisibleRow(m, s) || c != firstVisibleColumn(m, s)) continue;
                    XSSFCell cell = cell(m.getFirstRow(), m.getFirstColumn());
                    if (cell == null || isEmpty(cell)) continue;
                    try {
                        paintCellText(cv, s, cell, m);
                    } catch (RuntimeException e) {
                        if (cv != null) cv.restoreTo(depth);
                        // 1 つの壊れたセルで変換を止めない
                        Warnings.add("セル " + cell.getReference() + " の文字を描けませんでした", e);
                    }
                    continue;
                }
                if (L.cw(c) <= 0) continue;
                XSSFCell cell = row.getCell(c);
                if (cell == null || isEmpty(cell)) continue;
                try {
                    paintCellText(cv, s, cell, null);
                } catch (RuntimeException e) {
                    if (cv != null) cv.restoreTo(depth);
                    // 1 つの壊れたセルで変換を止めない
                    Warnings.add("セル " + cell.getReference() + " の文字を描けませんでした", e);
                }
            }
            if (s.c0 > 0 && s.body && !measuringRange) {
                // 前のページの列にある右へ傾けた文字の続き
                for (var cl : row) {
                    int c = cl.getColumnIndex();
                    XSSFCell xc = (XSSFCell) cl;
                    if (c >= s.c0 || isEmpty(xc) || L.mergedAt(r, c) != null) continue;
                    XSSFCellStyle cst = xc.getCellStyle();
                    if (cst == null) continue;
                    double sl = slantAngle(cst);
                    if (sl > 0 && L.x(c) + L.cw(c) + L.rh(r) / Math.tan(Math.toRadians(sl)) > L.x(s.c0)) {
                        try {
                            paintCellText(cv, s, xc, null);
                        } catch (RuntimeException e) {
                            if (cv != null) cv.restoreTo(depth);
                            Warnings.add("セル " + xc.getReference() + " の文字を描けませんでした", e);
                        }
                    }
                }
            }
            if (s.body && !measuringRange) {
                // 次のページの列にある左へ傾けた文字の、このページにかかる分
                for (var cl : row) {
                    int c = cl.getColumnIndex();
                    XSSFCell xc = (XSSFCell) cl;
                    if (c <= s.c1 || isEmpty(xc) || L.mergedAt(r, c) != null) continue;
                    XSSFCellStyle cst = xc.getCellStyle();
                    if (cst == null || !hasAnyBorder(cst)) continue;
                    double sl = slantAngle(cst);
                    if (sl < 0 && L.x(c) - L.rh(r) / Math.tan(Math.toRadians(-sl)) < L.x(s.c1 + 1)) {
                        try {
                            paintCellText(cv, s, xc, null);
                        } catch (RuntimeException e) {
                            if (cv != null) cv.restoreTo(depth);
                            Warnings.add("セル " + xc.getReference() + " の文字を描けませんでした", e);
                        }
                    }
                }
            }
            if (s.c0 > 0 && s.body && !measuringRange) {
                int k = spillColumn(row, r, s);
                // 印刷範囲のあるシート: 行のセルの定義がこのページの列まで届かない行は、Excel は続きを描かない
                if (k >= 0 && !ps.printAreas.isEmpty() && row.getLastCellNum() - 1 < s.c0) k = -1;
                if (k >= 0 && !ps.printAreas.isEmpty() && !screenSpill(row.getCell(k), s.c0)) k = -1;
                if (k >= 0 && row.getCell(k) != null && row.getCell(k).getCellStyle() != null && slantAngle(row.getCell(k).getCellStyle()) != 0) k = -1; // 斜めの文字は上で描いた
                if (k >= 0) {
                    try {
                        paintCellText(cv, s, row.getCell(k), null);
                    } catch (RuntimeException e) {
                        if (cv != null) cv.restoreTo(depth);
                        // 1 つの壊れたセルで変換を止めない
                        Warnings.add("セル " + row.getCell(k).getReference() + " の文字を描けませんでした", e);
                    }
                }
            }
        }
    }

    /** 結合範囲のうち、このブロックで最初に見える行 (先頭の行が非表示でも、残りの行で描く)。 */
    private int firstVisibleRow(CellRangeAddress m, Paginator.Segment s) {
        int last = Math.min(m.getLastRow(), s.r1);
        for (int r = Math.max(m.getFirstRow(), s.r0); r <= last; r++) if (L.rh(r) > 0) return r;
        return -1;
    }

    /** 結合範囲のうち、このブロックで最初に見える列。 */
    private int firstVisibleColumn(CellRangeAddress m, Paginator.Segment s) {
        int last = Math.min(m.getLastColumn(), s.c1);
        for (int c = Math.max(m.getFirstColumn(), s.c0); c <= last; c++) if (L.cw(c) > 0) return c;
        return -1;
    }

    /** 画面の寸法でも、セルの文字が toCol 列まではみ出すか (左揃えの文字列だけ判定する。それ以外は届くとみなす)。 */
    private boolean screenSpill(XSSFCell cell, int toCol) {
        try {
            if (cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) return true;
            XSSFCellStyle st = cell.getCellStyle();
            var al = st.getAlignment();
            if (al != org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL && al != org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT) return true;
            if (st.getRotation() != 0 || st.getIndention() != 0) return true;
            return L.screenReach(cell, cell.getStringCellValue(), toCol);
        } catch (RuntimeException e) {
            return true;
        }
    }


    /**
     * 行 r で、前のページのセルから文字がはみ出してこのブロックに届く列を返す (なければ -1)。
     *
     * @param s ブロック
     */
    private int spillColumn(XSSFRow row, int r, Paginator.Segment s) throws IOException {
        for (int k = Math.min(s.c0 - 1, row.getLastCellNum() - 1); k >= 0; k--) {
            if (L.mergedAt(r, k) != null) return -1;
            XSSFCell cell = row.getCell(k);
            if (cell == null || isEmpty(cell)) continue;
            boolean keep = measuringRange;
            measuringRange = true;
            rangeLast = k;
            try {
                paintCellText(null, new Paginator.Segment(r, r, k, s.c1, 0, 0, true), cell, null);
            } catch (RuntimeException e) {
                return -1;
            } finally {
                measuringRange = keep;
            }
            return rangeLast >= s.c0 ? k : -1;
        }
        return -1;
    }

    /** 前のページの文字のはみ出しがこのブロックに届くか (白紙の判定)。 */
    public boolean hasSpill(Paginator.Segment s) {
        // 印刷範囲のあるシートは何もないページも出力するので、調べなくてよい
        if (s.c0 <= 0 || !ps.printAreas.isEmpty()) return false;
        for (int r = s.r0; r <= s.r1; r++) {
            ctx.checkDeadline();
            if (L.rh(r) <= 0) continue;
            XSSFRow row = sheet.getRow(r);
            if (row == null) continue;
            try {
                if (spillColumn(row, r, s) >= 0) return true;
            } catch (IOException e) {
                return false;
            }
        }
        return false;
    }

    /**
     * 1 つのセル (結合範囲) の文字を描く: 配置、折り返し、縮小、回転、縦書き、はみ出し、収まらない数値の #### 表示。
     *
     * @param m 結合範囲 (なければ null)
     */
    private void paintCellText(PdfCanvas cv, Paginator.Segment s, XSSFCell cell, CellRangeAddress m) throws IOException {
        if (cv == null || !PdfCanvas.MARK) { paintCellText0(cv, s, cell, m); return; }
        // -Dx2p.mark=true: セルの属性をタグ名にして囲む (精度の内訳を測る開発用)
        XSSFCellStyle st = cell.getCellStyle();
        Align a = align(st);
        XSSFFont f = st.getFont();
        boolean rich = cell.getCellType() == CellType.STRING && cell.getRichStringCellValue().numFormattingRuns() > 1;
        String tag = "X2PC_" + a.h() + "_" + a.v() + "_w" + (a.wrap() ? 1 : 0) + "_r" + a.rotation() + "_i" + a.indent() + "_s" + (a.shrink() ? 1 : 0)
                + "_m" + (m == null ? "0" : (m.getLastRow() - m.getFirstRow() + 1) + "x" + (m.getLastColumn() - m.getFirstColumn() + 1))
                + "_t" + cell.getCellType() + "_R" + (rich ? 1 : 0) + "_z" + f.getFontHeightInPoints()
                + "_H" + (cell.getRow().getCTRow().isSetCustomHeight() && cell.getRow().getCTRow().getCustomHeight() ? 1 : 0)
                + "_F" + String.valueOf(f.getFontName()).replaceAll("[^\\p{L}\\p{N}]", "");
        cv.beginMark(tag);
        try { paintCellText0(cv, s, cell, m); } finally { cv.endMark(); }
    }

    private void paintCellText0(PdfCanvas cv, Paginator.Segment s, XSSFCell cell, CellRangeAddress m) throws IOException {
        int r = cell.getRowIndex(), c = cell.getColumnIndex();
        // 前のセルのアイコンの待ちを持ち越さない (はみ出しを調べる段階は描く前に抜けるので残る)
        pendingIcon = null;
        suppressText = false;
        XSSFCellStyle st = cell.getCellStyle();
        CondFormat.Result cfr = cf.eval(r, c);
        ValueFormatter.Result vr;
        if (L.showFormulas && cell.getCellType() == CellType.FORMULA) {
            // 数式の表示のシート: 数式をそのまま文字として印刷する
            vr = new ValueFormatter.Result();
            vr.text = "=" + cell.getCellFormula();
        } else if (L.showFormulas && cell.getCellType() == CellType.NUMERIC) {
            // 表示形式を使わず、数式バーと同じ値 (有効数字 15 桁、指数にしない) を出す
            vr = new ValueFormatter.Result();
            vr.text = new java.math.BigDecimal(cell.getNumericCellValue()).round(new java.math.MathContext(15))
                    .stripTrailingZeros().toPlainString();
        } else if (L.showFormulas && cell.getCellType() == CellType.BOOLEAN) {
            vr = new ValueFormatter.Result();
            vr.text = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
        } else {
            vr = ctx.formatter.format(cell, ps.showZeros, ps.errors, cfr != null ? cfr.numFmt : null);
        }
        // 表示形式で文字が空になるセル (;;; など) も、アイコンセットのアイコンは描く (Excel の「アイコンだけ」の作り方)
        boolean emptyIcon = vr.text.isEmpty() && cfr != null && cfr.icon != null && !bw() && cv != null;
        if (vr.text.isEmpty() && !emptyIcon) return;
        if (emptyIcon) vr.text = "0";
        // データバーの「棒のみ表示」は文字を描かない (同じセルのアイコンは描く)
        if (cfr != null && cfr.hideText && (cfr.icon == null || bw())) return;
        XSSFFont font = st.getFont();
        padFont = font;
        stackedText = align(st).rotation() == 255;
        double x0, y0, w, h;
        if (m != null) {
            x0 = L.x(m.getFirstColumn()); y0 = L.y(m.getFirstRow());
            w = L.x(m.getLastColumn() + 1) - x0; h = L.y(m.getLastRow() + 1) - y0;
        } else {
            x0 = L.x(c); y0 = L.y(r); w = L.cw(c); h = L.rh(r);
        }
        if (w <= 0 || h <= 0) return;
        // アイコンセット: アイコン (フォントの大きさ + 1pt の正方形) をセルの左に置き、文字はその右の範囲に配置する
        // 白黒印刷ではアイコンを描かず、場所も空けない
        IconSets.Icon icon = cfr != null && !bw() ? cfr.icon : null;
        double iconLeft = x0;
        if (icon != null) {
            double size = font.getFontHeight() / 20.0 + 1;
            double off = ICON_OFFSET * size;
            pendingIcon = icon;
            iconX = x0 + off;
            iconSize = size;
            x0 += off + size;
            w = Math.max(0, w - off - size);
            suppressText = icon.iconOnly() || emptyIcon || cfr.hideText;
        }

        Align al = align(st);
        // はみ出しを調べる段階 (cv が null、または measuringOverflow) では画像・チェックボックスを描かない (はみ出さず、2 回描かない)
        if (cell.getCTCell().isSetVm() && (cv == null || measuringOverflow || paintCellImage(cv, cell, al, x0, y0, w, h))) return;
        if (vr.bool && isCheckboxCell(st)) {
            if (cv == null || measuringOverflow) return;
            paintCellCheckbox(cv, cell, font, al, x0, y0, w, h);
            return;
        }
        spaceEnds = al.spaceEnds();
        mergedText = m != null;
        textBottomRow = m != null ? m.getLastRow() : r;
        cellWrap = align(st).wrap();
        HorizontalAlignment ha = al.h();
        VerticalAlignment va = al.v();
        // 数式の表示のシートでは、数値・真偽値・数式はセルの配置によらず左揃え (文字のセルは配置どおり。exp_showf)
        if (L.showFormulas && cell.getCellType() != CellType.STRING && cell.getCellType() != CellType.BLANK) ha = HorizontalAlignment.LEFT;
        if (ha == HorizontalAlignment.GENERAL) {
            // 縦書き (textRotation=255) の標準は中央
            if (al.rotation() == 255) ha = HorizontalAlignment.CENTER;
            // 上向き 90 度の文字の標準は右
            else if (al.rotation() == 90 && !vr.number && !vr.bool) ha = HorizontalAlignment.RIGHT;
            // 下向きに傾けた文字 (91〜179) の標準も右
            else if (al.rotation() > 90 && al.rotation() < 180 && !vr.number && !vr.bool) ha = HorizontalAlignment.RIGHT;
            else ha = vr.number ? HorizontalAlignment.RIGHT : vr.bool ? HorizontalAlignment.CENTER : HorizontalAlignment.LEFT;
        }
        // 右から左のシート: 本体を左右反転して描くので、論理の配置の左右を入れ替える (反転後に Excel と同じ向き)。
        // 印刷の範囲の計算 (はみ出しの届く列) は入れ替えない (Excel は左から右と同じく右の列へ数える)
        if (ps.rtl && !measuringRange) ha = ha == HorizontalAlignment.LEFT ? HorizontalAlignment.RIGHT : ha == HorizontalAlignment.RIGHT ? HorizontalAlignment.LEFT : ha;
        boolean wrap = al.wrap() || ha == HorizontalAlignment.JUSTIFY || ha == HorizontalAlignment.DISTRIBUTED
                || va == VerticalAlignment.JUSTIFY || va == VerticalAlignment.DISTRIBUTED;
        cellWrapLayout = wrap;
        int rot = al.rotation();
        boolean vertical = rot == 255;
        double angle = 0;
        if (!vertical && rot != 0) angle = rot <= 90 ? rot : -(rot - 90);
        rotWrap = wrap && angle != 0;
        if (angle != 0 || vertical) wrap = wrap && vertical;
        int indent = al.indent();
        // 字下げ 1 段の幅。paper: 用紙上で一定 (印刷倍率で縮めない。標準フォントの空白 3 つ分。RenderContext.indentStepPt)
        // ただし倍率 50% 未満では用紙上で一定にせず倍率で縮める
        boolean paper = L.scale >= INDENT_PAPER_MIN_SCALE;
        double indentPt = paper ? indent * ctx.indentStepPt / L.scale : indent * ctx.indentStepPt;
        boolean shrink = al.shrink() && !wrap;
        // 繰り返し文字 (*) のある書式は、右揃えで折り返すセルでは繰り返さず、普通の文字として配置する
        // (右揃えは末尾の _ の空白を除いてセルの右に寄る)
        if (vr.fillChar != null && al.wrap() && ha == HorizontalAlignment.RIGHT) vr.fillChar = null;

        cellRich = vr.rich != null;
        runLook = tables.look(r, c);
        // ピボットの行ラベル: 展開ボタンの分 + 段ごとの字下げ (ボタンは印刷しないが、Excel は場所を空ける)
        if (runLook != null && runLook.pivotLabel && (ha == HorizontalAlignment.LEFT || ha == HorizontalAlignment.GENERAL)) {
            double px = TableStyles.PIVOT_BUTTON_PX + indent * TableStyles.PIVOT_LEVEL_PX;
            indentPt = paper ? px * SheetLayout.PX / L.scale : px * SheetLayout.PX;
            if (ha == HorizontalAlignment.GENERAL) ha = HorizontalAlignment.LEFT;
        }
        List<TextLayout.Run> runs = runs(vr, font, cfr);
        runLook = null;
        // 末尾の改行は行として数えない
        if (trimNewlines(va, runs)) trimTrailingNewlines(runs);
        if (cellRich && !wrap) trimTrailingSpaces(runs);
        double padL = padX(), padR = padX();
        if (ha == HorizontalAlignment.LEFT || ha == HorizontalAlignment.DISTRIBUTED) padL += indentPt;
        if (ha == HorizontalAlignment.RIGHT || ha == HorizontalAlignment.DISTRIBUTED) padR += indentPt;
        double availW = Math.max(0, w - padL - padR);

        // 収まらない数値: 標準の表示形式なら桁を減らし、それでも収まらなければ # で埋める。真偽値も # で埋める
        if ((vr.number || vr.logical) && angle == 0 && !vertical && !shrink) {
            float tw = totalWidth(runs);
            // 標準の表示形式の数値は、どの文字も数字 1 文字分の幅として、列に入る文字数で決める
            int genMax = -1;
            // 長い数値に合わせて広げた列 (numWiden) は、Excel が数値の入る幅にしたので数えない
            if (vr.general && vr.number && !L.numWidened(c)) {
                float dw = runs.get(0).spec.width("0", runs.get(0).size, ctx.fonts);
                genMax = dw > 0 ? (int) Math.floor((availW + 0.01) / dw) : 0;
                if (vr.text.length() > genMax) tw = Float.MAX_VALUE;
            }
            // 日付として表せない値 (負、9999 年より後) は、幅によらずセルの幅いっぱいの #
            boolean badDate = vr.number && !vr.text.isEmpty() && vr.text.chars().allMatch(ch -> ch == '#');
            if (tw > availW + 0.01 || badDate) {
                String fitted = null;
                if (vr.general && genMax >= 0) fitted = ValueFormatter.generalFit(vr.value, genMax);
                else if (vr.general) {
                    for (int n = vr.text.length() - 1; n >= 1; n--) {
                        String t = ValueFormatter.general(vr.value, n);
                        if (runs.get(0).spec.width(t, runs.get(0).size, ctx.fonts) <= availW + 0.01) { fitted = t; break; }
                    }
                }
                if (fitted == null) {
                    // Excel は # をセルの配置によらず中央揃えで置き、個数は左右の余白に合わせて 2 画素はみ出してよい
                    float hw = runs.get(0).spec.width("#", runs.get(0).size, ctx.fonts);
                    // セルが # 1 つより狭くても、Excel は # を 1 つ描く
                    // 会計・円などの繰り返し文字 (*) のある書式は、# が 1 つも入らない狭い列では何も描かない (Excel の動き)
                    boolean repeatFmt = vr.fillChar != null;
                    int n = hw > 0 ? (int) Math.floor((availW + HASH_OVER_PX * 0.12 / Math.max(0.1, L.scale) + 1e-6) / hw) : 0;
                    if (!repeatFmt) n = Math.max(1, n);
                    fitted = n > 0 ? fractionHashes(vr.text, n) : "";
                    vr.fillChar = null;
                    ha = HorizontalAlignment.CENTER;
                }
                TextLayout.Run base = runs.get(0);
                runs = new ArrayList<>(List.of(base.copy(fitted)));
            }
        }

        // はみ出しに使える横の範囲
        double clipL = x0, clipR = x0 + w;
        boolean canOverflow = m == null && !wrap && !vr.number && !vr.bool && angle == 0 && !vertical && !shrink
                && (ha == HorizontalAlignment.LEFT || ha == HorizontalAlignment.RIGHT || ha == HorizontalAlignment.CENTER);
        if (ha == HorizontalAlignment.CENTER_SELECTION && m == null) {
            // 同じ配置の、続く空のセルにまたがって中央揃え
            int cc = c + 1;
            // 範囲は印刷範囲の外の列まで続く
            while (cc <= (!L.model().authorEnv ? L.maxCol - 1 : s.c1) && isEmpty(cell(r, cc)) && L.mergedAt(r, cc) == null) {
                XSSFCellStyle ns = style(r, cc);
                if (ns == null || align(ns).h() != HorizontalAlignment.CENTER_SELECTION) break;
                cc++;
            }
            w = L.x(cc) - x0;
            availW = Math.max(0, w - padL - padR);
            clipR = x0 + w;
            ha = HorizontalAlignment.CENTER;
            canOverflow = !wrap && !vr.number;
        }
        float textW = totalWidth(runs);
        if (canOverflow && textW > availW) {
            double need = textW + padL + padR;
            if (ha == HorizontalAlignment.LEFT) {
                clipR = extendRight(r, c, s, x0 + need);
            } else if (ha == HorizontalAlignment.RIGHT) {
                clipL = extendLeft(r, c, s, x0 + w - need);
            } else {
                double half = (need - w) / 2;
                clipR = extendRight(r, c, s, x0 + w + half);
                clipL = extendLeft(r, c, s, x0 - half);
            }
        }

        if (measuringRange && m == null && angle > 0 && angle < 90 && st != null && hasAnyBorder(st)) {
            // 右へ傾けた文字: 文字の右端がかかる列まで印刷する
            double mx = paintSlanted(null, TextLayout.layout(runs, 0, false, true, ctx.fonts), x0, y0, w, h, ha, va, angle);
            int cc = c;
            // 右の列に値のあるセルがあれば、そこで止める (普通のはみ出しと同じ。C 列に値があり Excel は広げない)
            while (cc + 1 <= s.c1 && L.x(cc + 1) < mx && isEmpty(cellAt(r, cc + 1))) cc++;
            rangeLast = Math.max(rangeLast, cc);
            return;
        }
        if (measuringRange) {
            if (m == null && !wrap && !vr.number && !vr.bool && angle == 0 && !vertical && !shrink
                    && (ha == HorizontalAlignment.LEFT || ha == HorizontalAlignment.CENTER)) measureRange(r, c, s, ha, runs, padL);
            return;
        }
        if (measuringOverflow) {
            // はみ出した範囲の内側の列の境目を記録する (Excel ははみ出した文字の下の縦罫線を描かない。
            // 余白だけがはみ出す場合は罫線を残す
            if (clipL < x0 - 0.01 || clipR > x0 + w + 0.01) {
                double t0 = switch (ha) {
                    case LEFT -> x0 + padL;
                    case RIGHT -> x0 + w - padR - textW;
                    default -> x0 + (w - textW) / 2;
                };
                double t1 = Math.min(clipR, t0 + textW);
                t0 = Math.max(clipL, t0);
                for (int cc = s.c0; cc <= s.c1 + 1; cc++) {
                    double xe = L.x(cc);
                    if (xe > t0 + 0.01 && xe < t1 - 0.01 && (xe < x0 + 0.01 || xe > x0 + w - 0.01)) overflowEdges.add(r * 100000L + cc);
                }
            }
            return;
        }
        float sizeScale = 1;
        if (vertical && shrink) {
            // 縦書きは、積んだ高さをセルの高さに収める
            // 改行で分かれた列ごとに積んだ高さを求め、いちばん高い列を収める
            int n = 0;
            double need = 0, col = 0;
            for (TextLayout.Run rr : runs) {
                double pitch = stackPitch(rr);
                for (int cp : rr.text.codePoints().toArray()) {
                    if (cp == '\r') continue;
                    if (cp == '\n') { need = Math.max(need, col); col = 0; continue; }
                    n++;
                    col += pitch;
                }
            }
            need = Math.max(need, col);
            double avail = h - padTop() - padBottom();
            if (n > 0 && need > avail) sizeScale = (float) (avail / need);
        } else if (shrink && textW > availW && textW > 0) {
            double[] pad = shrinkSteps(runs, w, padL - padX(), padR - padX());
            padL = pad[0];
            padR = pad[1];
            availW = Math.max(0, w - padL - padR);
        }
        if (sizeScale < 1) {
            // 横書きの縮小と同じく 1 ppem (600 dpi の 1 画素) より小さくしない (余白より低いセルで倍率が 0 以下にならないように)
            float min = (float) (0.12 / Math.max(0.1, L.scale));
            for (TextLayout.Run rr : runs) rr.size = Math.max(min, rr.size * sizeScale);
        }

        // 結合セルが改ページをまたぐと、文字はそのブロックの範囲で切り抜かれる
        double clipT = y0, clipB = y0 + h;
        if (posx() && s != null) {
            clipT = Math.max(clipT, L.y(s.r0));
            clipB = Math.min(clipB, L.y(s.r1 + 1));
        }
        if (icon != null) clipL = Math.min(clipL, iconLeft);
        if (angle != 0 && Math.abs(angle) < 90) {
            // 斜めの文字は平行四辺形の外枠まで描く
            double sh = h / Math.tan(Math.toRadians(Math.abs(angle))) * Math.signum(angle);
            clipL = Math.min(clipL, x0 + Math.min(0, sh));
            clipR = Math.max(clipR, x0 + w + Math.max(0, sh));
        }
        cv.save();
        try {
        cv.clipRect(clipL, clipT, clipR - clipL, Math.max(0, clipB - clipT));
        if (vertical) {
            if (!suppressText) paintStacked(cv, runs, x0, y0, w, h, ha, va, indentPt, wrap);   // アイコンだけのときは文字を描かない
        } else if (angle != 0) {
            slantBox = m == null && st != null && hasAnyBorder(st);
            if (!suppressText) paintRotated(cv, runs, x0, y0, w, h, ha, va, angle, padL, padR);
        } else if (vr.fillChar != null && vr.fillLeft != null) {
            paintAccounting(cv, runs.get(0), vr, x0 + padL, y0, availW, h, va);
        } else if (ha == HorizontalAlignment.FILL) {
            TextLayout.Run r0 = runs.get(0);
            float one = totalWidth(runs);
            int n = one > 0 ? Math.max(1, (int) Math.floor(availW / one)) : 1;
            List<TextLayout.Run> rep = new ArrayList<>();
            for (int i = 0; i < n; i++) for (TextLayout.Run rr : runs) rep.add(rr.copy(rr.text));
            List<TextLayout.Line> lines = TextLayout.layout(rep, 0, false, false, ctx.fonts);
            paintLines(cv, lines, x0, y0, w, h, HorizontalAlignment.LEFT, va, padL, padR, r0);
        } else {
            // 会計用の下線: 文字列のセルはセルの幅いっぱい (数値は文字の幅)
            if (!vr.number && runs.size() == 1 && runs.get(0).accounting && !wrap) {
                runs.get(0).ulFrom = x0 + padL;
                runs.get(0).ulTo = x0 + w - padR;
            }
            // 折り返しの右の余白は 2 画素小さい (Excel の折り返した行は使える幅を 2 画素まで超える。右揃えの位置も同じ)
            double wrapW = posx() && cellWrap && wrap ? availW + 2 * 0.12 / Math.max(0.1, L.scale) : availW;
            List<TextLayout.Line> lines = TextLayout.layout(runs, (float) wrapW, wrap, wrap, ctx.fonts);
            newlineHeights(lines);
            phon = vr.fillChar == null ? phonetic(cell, cfr) : null;
            try {
                paintLines(cv, lines, x0, y0, w, h, ha, va, padL, padR, runs.get(0));
            } finally {
                phon = null;
            }
        }
            // 縦書き・斜めの文字の経路はアイコンを描かないので、ここで描く (最初のランの基線をセルの縦位置で決める)
            if (pendingIcon != null) {
                TextLayout.Run r0 = runs.get(0);
                double asc = r0.spec.ascent * r0.size, desc = r0.spec.descent * r0.size;
                double baseline = switch (va) {
                    case TOP -> y0 + padTop() + asc;
                    case CENTER -> y0 + (h - asc - desc) / 2 + asc + centerDy();
                    default -> y0 + h - padBottom() - desc;
                };
                paintPendingIcon(cv, baseline, r0.size);
            }
        } finally {
            pendingIcon = null;
            suppressText = false;
            cv.restore();
        }
    }

    /** アイコンの左の余白 (アイコンの大きさに対する割合。Excel の印刷: 12pt のアイコンで 1.56pt)。 */
    static final double ICON_OFFSET = 0.13;
    /** アイコンの下端の、文字の基線からの距離 (フォントの大きさに対する割合)。 */
    static final double ICON_BELOW_BASELINE = 0.3;
    /** 描く前のアイコンセットのアイコン (最初の行の基線が決まったところで描く)。 */
    private IconSets.Icon pendingIcon;
    private double iconX, iconSize;
    /** 値を表示しない (アイコンのみ表示)。 */
    private boolean suppressText;

    /** 待っているアイコンを、基線に合わせて描く。 */
    private void paintPendingIcon(PdfCanvas cv, double baseline, float fontSize) throws IOException {
        if (pendingIcon == null) return;
        double bottom = baseline + ICON_BELOW_BASELINE * fontSize;
        IconPainter.paint(cv, pendingIcon.set(), pendingIcon.index(), iconX, bottom - iconSize, iconSize);
        pendingIcon = null;
    }

    /** 右の空のセルへ、needX まで (ブロック内で) はみ出せる右端。 */
    private double extendRight(int r, int c, Paginator.Segment s, double needX) {
        int cc = c + 1;
        double x = L.x(cc);
        while (x < needX && cc <= s.c1) {
            if (!isEmpty(cell(r, cc)) || L.mergedAt(r, cc) != null) break;
            x += L.cw(cc);
            cc++;
        }
        return x;
    }

    /** 左の空のセルへ、needX まで (ブロック内で) はみ出せる左端。 */
    private double extendLeft(int r, int c, Paginator.Segment s, double needX) {
        int cc = c - 1;
        double x = L.x(c);
        while (x > needX && cc >= s.c0) {
            if (!isEmpty(cell(r, cc)) || L.mergedAt(r, cc) != null) break;
            x -= L.cw(cc);
            cc--;
        }
        return x;
    }

    /** 末尾の改行を行として数えずに取り除くか。 */
    private static boolean trimNewlines(VerticalAlignment va, List<TextLayout.Run> runs) {
        return va == VerticalAlignment.DISTRIBUTED;
    }

    /** 最後のランから順に、末尾の半角・全角の空白を取り除く (空になったランは除く。少なくとも 1 つのランは残す)。 */
    private static void trimTrailingSpaces(List<TextLayout.Run> runs) {
        while (!runs.isEmpty()) {
            TextLayout.Run last = runs.get(runs.size() - 1);
            String t = last.text;
            int e = t.length();
            while (e > 0 && (t.charAt(e - 1) == ' ' || t.charAt(e - 1) == '\u3000')) e--;
            if (e == 0 && runs.size() > 1) { runs.remove(runs.size() - 1); continue; }
            last.text = t.substring(0, e);
            return;
        }
    }

    /** 最後のランの末尾の改行 (\r, \n) を取り除く (空になったランは除く。少なくとも 1 つのランは残す)。 */
    private static void trimTrailingNewlines(List<TextLayout.Run> runs) {
        while (!runs.isEmpty()) {
            TextLayout.Run last = runs.get(runs.size() - 1);
            String t = last.text;
            int e = t.length();
            while (e > 0 && (t.charAt(e - 1) == '\n' || t.charAt(e - 1) == '\r')) e--;
            if (e == t.length()) return;
            if (e == 0 && runs.size() > 1) { runs.remove(runs.size() - 1); continue; }
            last.text = t.substring(0, e);
            return;
        }
    }

    /** ランの幅の合計 (改行文字を除く)。 */
    private float totalWidth(List<TextLayout.Run> runs) {
        float w = 0;
        for (TextLayout.Run r : runs) {
            String t = r.text.replace("\n", "").replace("\r", "");
            TextLayout.Run x = r.copy(t);
            w += TextLayout.width(x, ctx.fonts);
        }
        return w;
    }

    /**
     * Excel が折り返した行に使う行送り: 行の最大のフォントをプリンタ解像度で実体化したときの自動調整行の高さ
     * (h + floor((h+4)/10) + 1 デバイス px)。折り返した行の組 約 9000 で実測。
     */
    private double pitch(TextLayout.Line l) {
        if (l.runs.isEmpty() && l.newlineRun != null && !L.model().authorEnv) {
            double bp = L.linePitch(l.newlineRun.spec.requested, l.newlineRun.size);
            if (bp > 0) return bp;
        }
        // 空行の送りも、その行のフォントの行送り
        if (l.runs.isEmpty() && l.blankBase != null && !L.model().authorEnv) {
            double bp = L.linePitch(l.blankBase.spec.requested, l.blankBase.size);
            if (bp > 0) return bp;
        }
        if (l.runs.isEmpty()) return (l.ascent + l.descent) * LINE_FACTOR;
        TextLayout.Run big = bigRun(l);
        double p = L.linePitch(big.spec.requested, big.size);
        // 書式付き文字列の行送りは、小さいランだけの行でもセルのフォントの行送り以上
        return p > 0 ? p : (l.ascent + l.descent) * LINE_FACTOR;
    }

    /** セルのふりがな (文字列の文字位置 sb〜eb の上に t を描く)。 */
    private static final class Phonetic {
        String text;
        int[] sb, eb;
        String[] t;
        TextLayout.Run font;
        /** 各行の上に確保する高さ (pt)。 */
        double lineH;
        /** 配置 (left、center、distributed、noControl)。 */
        String align = "left";
    }

    /** ふりがなの基線の、ふりがなの行の下端からの下がり (ふりがなの文字の大きさの倍。Excel の印刷の実測)。 */
    static final double RUBY_DROP = 0.1;

    /** 描いているセルのふりがな (なければ null)。{@link #paintLines} が使う。 */
    private Phonetic phon;

    /**
     * 「ふりがなを表示」(セルの ph="1") の文字列のふりがな。Excel は各行の上にふりがなのフォント 1 行分の高さを足し、
     * 対象の文字の上に phoneticPr のフォントで描く。
     */
    private Phonetic phonetic(XSSFCell cell, CondFormat.Result cfr) {
        if (cell.getCellType() != CellType.STRING) return null;
        var cc = cell.getCTCell();
        if (!cc.isSetPh() || !cc.getPh()) return null;
        XSSFRichTextString rs = cell.getRichStringCellValue();
        var ct = rs == null ? null : rs.getCTRst();
        if (ct == null || ct.sizeOfRPhArray() == 0) return null;
        XSSFFont pf = null;
        if (ct.isSetPhoneticPr()) {
            try {
                pf = cell.getSheet().getWorkbook().getStylesSource().getFontAt((int) ct.getPhoneticPr().getFontId());
            } catch (RuntimeException ignored) {
                pf = null;
            }
        }
        if (pf == null) pf = cell.getCellStyle().getFont();
        Phonetic p = new Phonetic();
        p.text = rs.getString();
        p.font = run("", pf, null, null);
        // ふりがなの行の高さ: 字の高さ floor((ascent + descent) × 画面の字の大きさ) (行の自動の高さと同じ。標準行の行間を含まない)
        long ppem = Math.round(pf.getFontHeight() / 20.0 * 4 / 3.0);
        p.lineH = Math.floor((p.font.spec.ascent + p.font.spec.descent) * ppem + 1e-6) * SheetLayout.PX;
        p.align = ct.isSetPhoneticPr() && ct.getPhoneticPr().isSetAlignment() ? ct.getPhoneticPr().getAlignment().toString() : "left";
        int n = ct.sizeOfRPhArray();
        p.sb = new int[n]; p.eb = new int[n]; p.t = new String[n];
        for (int i = 0; i < n; i++) {
            var ph = ct.getRPhArray(i);
            p.sb[i] = (int) ph.getSb();
            p.eb[i] = (int) ph.getEb();
            p.t[i] = ph.getT() == null ? "" : ph.getT();
        }
        return p;
    }

    /** 行 l の先頭から、行の文字列の n 文字 (char) 目までの幅。 */
    private float prefixWidth(TextLayout.Line l, int n) {
        float w = 0;
        for (TextLayout.Run r : l.runs) {
            if (n <= 0) break;
            int k = Math.min(n, r.text.length());
            w += TextLayout.width(r.copy(r.text.substring(0, k)), ctx.fonts);
            n -= k;
        }
        return w;
    }

    /** 行 l のふりがなを描く。lineStart は行の文字列が元の文字列で始まる位置。 */
    private void paintPhonetic(PdfCanvas cv, TextLayout.Line l, int lineStart, int len, double x, double rubyBaseline) throws IOException {
        TextLayout.Run f0 = phon.font;
        if ("noControl".equals(phon.align)) {
            // 指定なし: ふりがなを続けて、行の最初の対象の文字の位置から並べる
            StringBuilder all = new StringBuilder();
            int first = -1;
            for (int k = 0; k < phon.sb.length; k++) {
                if (phon.sb[k] < lineStart || phon.sb[k] >= lineStart + len) continue;
                if (first < 0) first = phon.sb[k];
                all.append(phon.t[k]);
            }
            if (first >= 0 && all.length() > 0) cv.text(all.toString(), x + prefixWidth(l, first - lineStart), rubyBaseline, f0.spec, f0.size, f0.color, 1, 0);
            return;
        }
        for (int k = 0; k < phon.sb.length; k++) {
            int sb = phon.sb[k], eb = Math.max(phon.sb[k] + 1, phon.eb[k]);
            if (sb < lineStart || sb >= lineStart + len || phon.t[k].isEmpty()) continue;
            int e = Math.min(eb, lineStart + len);
            double bx = x + prefixWidth(l, sb - lineStart);
            double bw = prefixWidth(l, e - lineStart) - (bx - x);
            TextLayout.Run f = phon.font;
            double rw = f.spec.width(phon.t[k], f.size, ctx.fonts);
            if ("distributed".equals(phon.align) && rw < bw && phon.t[k].codePointCount(0, phon.t[k].length()) > 1) {
                // 均等割り付け: 文字の間と両端 (半分) に同じ空きを入れて、対象の文字の幅に広げる
                int n = phon.t[k].codePointCount(0, phon.t[k].length());
                double gap = (bw - rw) / n, cx = bx + gap / 2;
                for (int i = 0; i < phon.t[k].length(); ) {
                    int cp = phon.t[k].codePointAt(i);
                    String ch = new String(Character.toChars(cp));
                    cv.text(ch, cx, rubyBaseline, f.spec, f.size, f.color, 1, 0);
                    cx += f.spec.width(ch, f.size, ctx.fonts) + gap;
                    i += Character.charCount(cp);
                }
                continue;
            }
            double rx = "center".equals(phon.align) || "distributed".equals(phon.align) ? bx + (bw - rw) / 2 : bx;
            cv.text(phon.t[k], rx, rubyBaseline, f.spec, f.size, f.color, 1, 0);
        }
    }

    /** 折り返した最後の行の高さ (画素): round((100% の文字の高さ + 10) × 倍率)。 */
    private long lastLinePx(TextLayout.Line l, double u, double sc) {
        if (!l.runs.isEmpty()) {
            TextLayout.Run big = bigRun(l);
            int[] ad = FontCalc.ascentDescent(big.spec.requested, (int) Math.round(rawPt(big) * 600 / 72));
            if (ad != null) return Math.round((ad[0] + ad[1] + 10) * sc);
        }
        return Math.round(pitch(l) / u);
    }

    /**
     * 行の高さを決める最大のラン (文字のある行のみ)。Excel のセルでは行を終える改行文字の書式も数える (大きい
     * 改行文字の行は高くなり、空行の高さはその空行を終える改行文字の書式)。
     */
    private TextLayout.Run bigRun(TextLayout.Line l) {
        TextLayout.Run big = l.runs.get(0);
        for (TextLayout.Run r : l.runs) if (r.size > big.size) big = r;
        // 改行文字の書式は、行が高くなるときだけ (大きさではなく行送りで比べる。游ゴシック 10pt の行の HG丸ゴシック 12pt の改行は効かない)
        if (l.newlineRun != null && runPitch(l.newlineRun) > runPitch(big)) big = l.newlineRun;
        return big;
    }

    /** ランの行送り (比べる用。求められなければ大きさ)。 */
    private double runPitch(TextLayout.Run r) {
        double p = posx() ? L.linePitch(r.spec.requested, r.size) : 0;
        return p > 0 ? p : r.size * 1.2;
    }

    /**
     * 行のアセント・ディセントに、行を終える改行文字の書式を入れる (空行はその書式だけ。末尾の空行は前の改行の書式)。
     * {@link #bigRun} と同じ規則。
     */
    private void newlineHeights(List<TextLayout.Line> lines) {
        for (int i = 0; i < lines.size(); i++) {
            TextLayout.Line l = lines.get(i);
            TextLayout.Run n = l.newlineRun;
            // 末尾の改行の後の最後の空行は、前の行と同じ高さ (前の行の高さを決めるラン。前の行も空行ならその改行の書式)
            if (n == null && i > 0 && i == lines.size() - 1 && l.runs.isEmpty()) {
                TextLayout.Line p = lines.get(i - 1);
                n = p.runs.isEmpty() ? p.newlineRun : bigRun(p);
            }
            if (n == null) continue;
            float a = (float) (n.spec.ascent * n.size), d = (float) (n.spec.descent * n.size);
            if (l.runs.isEmpty()) { l.blankBase = n; l.ascent = a; l.descent = d; }
            else if (bigRun(l) == n) { l.ascent = Math.max(l.ascent, a); l.descent = Math.max(l.descent, d); }
        }
    }

    /** メイリオ・Meiryo UI・Yu Gothic UI (winAscent が大きく、typo の行間を持つ書体)。 */
    static boolean uiTallFont(String font) {
        if (font == null) return false;
        String n = FontManager.norm(font).replace(" ", "");
        return n.startsWith("メイリオ") || n.startsWith("meiryo") || n.startsWith("yugothicui");
    }

    /** ランの大きさ (実体化の前の pt。0.5pt 刻みに丸める)。 */
    private double rawPt(TextLayout.Run r) { return Math.round(r.size * 2) / 2.0; }

    /** 行に全角の文字があるか。 */
    private static boolean hasWide(TextLayout.Line l) {
        for (TextLayout.Run r : l.runs) {
            for (int i = 0; i < r.text.length(); i++) if (r.text.charAt(i) >= 0x2E80) return true;
        }
        return false;
    }

    /** 文字の始まりの丸めの加算。 */
    static final double X_SNAP_ADD = 0.45;

    /**
     * 上揃え・縦中央の 1 行目の上端 (シートの pt。これに行のアセントを足すと基線)。セルの上下は整数の画素に切り捨てる。H はセルの高さ、h は文字の高さ、A はアセント (画素)、L は行送り。
     * 折り返しなし: 上揃え A + 6、中央 floor((H − (h + 10)) / 2) + A + 6。折り返しあり: 上揃え A + min(6, L − h)、中央 trunc((H − n L) / 2) + A + min(6, L − h)。
     */
    private double topPx(List<TextLayout.Line> lines, double y0, double h, VerticalAlignment va) {
        double extraAscent;
        double sc = Math.max(0.01, L.scale), u = 0.12 / sc;
        TextLayout.Line l = lines.get(0);
        double org = blockOy - blockSy * sc;
        long t = (long) Math.floor((org + y0 * sc) / 0.12 + 1e-6), b = (long) Math.floor((org + (y0 + h) * sc) / 0.12 + 1e-6);
        long ht = b - t;
        long a = Math.round(l.ascent / u), d = Math.round(l.descent / u);
        if (!l.runs.isEmpty()) {
            TextLayout.Run big = bigRun(l);
            int[] ad = FontCalc.ascentDescent(big.spec.requested, (int) Math.round(big.size / u));
            if (ad != null) { a = ad[0]; d = ad[1]; }
        }
        long fh = a + d;
        extraAscent = (a - l.ascent / u) * u;
        int n = lines.size();
        long off;
        // 上付き・下付きの小さい文字が上下にはみ出す分 (画素)。Excel は行の高さに入れて中央に置く
        long ea = 0, eb = 0;
        for (TextLayout.Run r : l.runs) {
            if (r.baselineShift == 0) continue;
            double top = r.shift() + r.spec.ascent * r.effSize(), bot = -r.shift() + r.spec.descent * r.effSize();
            ea = Math.max(ea, Math.round(top / u) - a);
            eb = Math.max(eb, Math.round(bot / u) - d);
        }
        if (!cellWrap) {
            off = va == VerticalAlignment.TOP ? 6 + ea : Math.floorDiv(ht - (fh + ea + eb + 10), 2L) + 6 + ea;
        } else {
            long lp = Math.round(pitch(l) / u), total = 0;
            for (int i = 0; i + 1 < n; i++) total += Math.round(pitch(lines.get(i)) / u);
            total += lastLinePx(lines.get(n - 1), u, sc);
            lp = Math.min(lp, lastLinePx(l, u, sc));
            long tp = Math.min(6, Math.max(0, lp - fh));
            // 折り返しの基線の位置: round((100% のアセント + 2.4) × 倍率 + 4)
            if (!l.runs.isEmpty()) {
                TextLayout.Run big = bigRun(l);
                int[] a0 = FontCalc.ascentDescent(big.spec.requested, (int) Math.round(rawPt(big) * 600 / 72));
                if (a0 != null) tp = Math.round((a0[0] + 2.4) * sc + 4) - a;
            }
            off = va == VerticalAlignment.TOP ? tp : (ht - total) / 2 + tp;
        }
        if (va == VerticalAlignment.CENTER && mergedText) off -= (long) MERGE_CENTER_PX;
        // メイリオ・Meiryo UI・Yu Gothic UI の上・中央揃えは、Excel の方が上 (
        // 上揃えは折り返しなしで 2 画素、折り返しで round(2 × 倍率)。中央揃えは 2 + floor(2 × 倍率 + 0.3)、折り返しで 1 + round(2.5 × 倍率))
        if (!l.runs.isEmpty() && uiTallFont(bigRun(l).spec.requested)) {
            if (va == VerticalAlignment.TOP) off -= cellWrap ? Math.round(2 * sc) : 2;
            else off -= cellWrap ? 1 + Math.round(2.5 * sc) : 2 + (long) Math.floor(2 * sc + 0.3);
        }
        if (va == VerticalAlignment.TOP && !l.runs.isEmpty()) {
            TextLayout.Run big = bigRun(l);
            off += FontCalc.topExtraPx(big.spec.requested, (int) Math.round(big.size / u));
        }
        return ((t + off) * 0.12 - org) / sc + extraAscent;
    }

    /** 行を縦・横の配置に合わせて描く (縦の両端揃え・均等割り付けは行の間に配る)。 */
    private void paintLines(PdfCanvas cv, List<TextLayout.Line> lines, double x0, double y0, double w, double h,
                            HorizontalAlignment ha, VerticalAlignment va, double padL, double padR,
                            TextLayout.Run first) throws IOException {
        if (lines.isEmpty()) return;
        double total = 0;
        for (int i = 0; i < lines.size(); i++) {
            TextLayout.Line l = lines.get(i);
            if (l.ascent + l.descent == 0) { l.ascent = first.spec.ascent * first.size; l.descent = first.spec.descent * first.size; }
            total += i == lines.size() - 1 ? l.ascent + l.descent : pitch(l);
        }
        double rubyH = phon != null ? phon.lineH : 0;
        total += rubyH * lines.size();
        // 結合セルに 1 行分以上収まらない複数行の文字は、縦の配置によらず先頭の行から描く。
        // わずかなはみ出しは配置どおり (実ファイルの様式は行の見積もりの差で収まらない側に出ることがある)
        if (mergedText && lines.size() > 1 && total > h - padTop() - padBottom() + pitch(lines.get(0)))
            va = va == VerticalAlignment.JUSTIFY || va == VerticalAlignment.DISTRIBUTED ? VerticalAlignment.JUSTIFY : VerticalAlignment.TOP;
        // 縦の均等割り付けの 1 行は中央揃えと同じ位置 (結合セルで書体によらず中央揃えの補正の分 3 画素、単独で 1 画素の違いになる)
        if (va == VerticalAlignment.DISTRIBUTED && lines.size() == 1) va = VerticalAlignment.CENTER;
        int phonPos = 0;
        double top;
        double gap = 0;
        double innerH = h - padTop() - padBottom();
        switch (va) {
            case TOP -> top = y0 + padTop();
            case CENTER -> top = y0 + (h - total) / 2 + centerDy() - (mergedText && posx() ? MERGE_CENTER_PX * 0.12 / Math.max(0.1, L.scale) : 0);
            case JUSTIFY, DISTRIBUTED -> {
                // 2 行以上は、最後の行の基線が下端から ディセント + 12 画素
                if (posx() && lines.size() > 1) innerH -= 6.5 * 0.12 / Math.max(0.1, L.scale);
                if (lines.size() > 1 && total < innerH) {
                    gap = (innerH - total) / (lines.size() - 1);
                    top = y0 + padTop();
                } else if (va == VerticalAlignment.DISTRIBUTED) {
                    top = y0 + (h - total) / 2;
                } else {
                    top = y0 + padTop();
                }
            }
            default -> {
                top = y0 + h - padBottom() - total;
                // 下揃えの基線は、シートの標準行を決めるフォントのディセントにも左右される
                if (posx() && !lines.isEmpty()) {
                    double sd = Math.max(0, baseDescent(textBottomRow) - lines.get(lines.size() - 1).descent);
                    // 同じ行のディセントの大きい下揃えのセルの分 (行の最大と自分の書体を同じ求め方で比べる)
                    if (padFont != null)
                        sd = Math.max(sd, rowDescent(textBottomRow) - descentPt(padFont));
                    // ＭＳ 系の書体で標準フォントも ＭＳ 系のブック: Excel の下揃えのディセントは、標準フォントのディセント (画面 (96 dpi) の
                    // 整数の画素で丸めた値との差が 0.5 画素以内ならその値、ほかは丸めない値。標準 6・7.5・8pt のブックは丸めない) + 画面の画素で数えたディセントの標準フォントとの差 (
                    // 標準 11pt のブックは 12.5 画素、標準 8pt は 9.4 画素。14pt は標準 11pt より 1 画面画素大きい)。小さくなるときは下げる
                    // シートの標準行・同じ行のセルのディセントが、セルと標準フォントの大きい方を超える行と、行のディセントがほかの書体のラン
                    // (記号の書体など) で大きくなっている行は除く
                    double ld = lines.get(lines.size() - 1).descent;
                    if (padFont != null && PrinterModel.isMs(FontManager.norm(padFont.getFontName()))
                            && PrinterModel.isMs(FontManager.norm(ctx.defaultFontName)) && Math.abs(ld - descentPt(padFont)) < 0.02) {
                        double own = descentPt(padFont), std = ctx.fonts.spec(ctx.defaultFontName, false, false).descent * ctx.defaultFontSize;
                        double gov = Math.max(own, std);
                        if (baseDescent(textBottomRow) <= gov + 1e-9 && rowDescent(textBottomRow) <= gov + 1e-9) {
                            double stdScr = screenDescPt(ctx.defaultFontName, ctx.defaultFontSize);
                            sd = (Math.abs(stdScr - std) <= 0.06 ? stdScr : std) + Math.max(0, screenDescPt(padFont.getFontName(), padFont.getFontHeight() / 20.0) - stdScr) * SCR_STEP - ld;
                        }
                    }
                    // 結合セルの折り返し (均等割り付けなど折り返す配置を含む)・リッチテキストは、少なくとも MERGE_WRAP_BOTTOM_PX 上
                    // (標準フォントによらず 0.56〜0.88pt。リッチテキスト・均等割り付けは折り返しの設定なしでも同じ)
                    double sdDesc = sd;
                    if (mergedText && (cellWrapLayout || cellRich)) sd = Math.max(sd, MERGE_WRAP_BOTTOM_PX * 0.12 / Math.max(0.1, L.scale));
                    // 行に thickBot (太い下罫線の印) があると、Excel は下揃えの文字を 8 画素 (倍率で縮む。0.96pt × 倍率) 上に置く (
                    // 罫線の太さではなく行の印で決まる)。結合セルの折り返しの持ち上げとは足さずに大きい方。
                    // 標準フォントが ＭＳ 系のブックだけ (游ゴシックのブックでは合わない)
                    if (PrinterModel.isMs(FontManager.norm(ctx.defaultFontName)) && textBottomRow >= 0 && sheet.getRow(textBottomRow) != null && sheet.getRow(textBottomRow).getCTRow().isSetThickBot()
                            && sheet.getRow(textBottomRow).getCTRow().getThickBot()) sd = Math.max(sd, sdDesc + 8 * 0.12);
                    top -= sd;
                }
            }
        }
        if (posx() && !L.model().authorEnv && rubyH == 0 && (va == VerticalAlignment.CENTER || va == VerticalAlignment.TOP)) {
            top = topPx(lines, y0, h, va);
        }
        double y = top;
        double availW = w - padL - padR;
        double textDy = textDy();
        for (int i = 0; i < lines.size(); i++) {
            TextLayout.Line l = lines.get(i);
            double baseline = y + rubyH + l.ascent + textDy / L.scale;
            double x;
            float extra = 0;
            int mode = 0;
            switch (ha) {
                // 右揃えの折り返しは、折り返しなしより 2 画素右
                case RIGHT -> x = x0 + w - padR - l.width + (posx() && cellWrap ? 2 * 0.12 / Math.max(0.1, L.scale) : 0);
                // 中央揃えの位置: 既定は行末の空白も含めた幅 (CENTER_TRAIL_SPACE)
                case CENTER -> x = x0 + (w - (l.width)) / 2;
                case JUSTIFY -> {
                    x = x0 + padL;
                    if (!l.endsParagraph && posx() && TextLayout.distGaps(l) > 0 && hasWide(l)) {
                        // 全角の文字のある行は、字間にも余白を配分して右端まで揃える
                        extra = (float) Math.max(0, availW + 2 * 0.12 / Math.max(0.1, L.scale) - l.width);
                        mode = 2;
                    } else if (!l.endsParagraph && l.spaces > 0) { extra = (float) (availW - l.width); mode = 1; }
                }
                case DISTRIBUTED -> {
                    int dg = TextLayout.distGaps(l);
                    if (dg == 0) {
                        x = x0 + (w - l.width) / 2;
                    } else {
                        x = x0 + padL;
                        // 右端は使える幅の 2 画素先まで (書体・倍率によらない。両端揃えと同じ)
                        extra = (float) Math.max(0, availW + (posx() ? 2 * 0.12 / Math.max(0.1, L.scale) : 0) - l.width);
                        mode = 2;
                        if (spaceEnds) {
                            // 「前後にスペースを入れる」(justifyLastLine) の均等割り付けは、余りを (文字数 + 1) 等分して
                            // 文字の間と両端に置く。設定がなければ最初と最後の文字をセルの端に付ける
                            double slot = extra / (dg + 2);
                            x += slot;
                            extra = (float) (slot * dg);
                        }
                        if (availW < l.width) x = x0 + (w - l.width) / 2;
                    }
                }
                default -> x = x0 + padL;
            }
            if (ha == HorizontalAlignment.CENTER && newLayout()) x += (posx() && cellWrap ? POSM_CENTER - 3 * 0.12 : POSM_CENTER) / L.scale;
            // 文字の始まりは整数の画素 (奇数の線幅のシートではセルの境界に 0.5 画素の端数がある)
            if (posx()) {
                double sc = Math.max(0.01, L.scale), org = blockOx - blockSx * sc;
                if (ha == HorizontalAlignment.CENTER && newLayout() && sc != 1) {
                    // 中央揃え: セルの左右の端を先に画素に丸めてから、(幅 − 文字の幅) の半分を切り捨てる (縮小で境界が画素の途中に
                    // 来るとき。丸めないと倍率 1.0・0.5 以外で 1 画素右になる)
                    long pl = (long) Math.floor((org + x0 * sc) / 0.12 + X_SNAP_ADD), pr = (long) Math.floor((org + (x0 + w) * sc) / 0.12 + X_SNAP_ADD);
                    long tw = Math.round(l.width * sc / 0.12);
                    long off = Math.round((cellWrap ? POSM_CENTER - 3 * 0.12 : POSM_CENTER) / 0.12);
                    x = ((pl + Math.floorDiv(pr - pl - tw, 2L) + off) * 0.12 - org) / sc;
                } else
                x = (Math.floor((org + x * sc) / 0.12 + X_SNAP_ADD) * 0.12 - org) / sc;
            }
            if (i == 0) paintPendingIcon(cv, baseline, first.size);
            if (!suppressText) TextLayout.drawLine(cv, l, x, baseline, extra, mode);
            if (phon != null) {
                StringBuilder lt = new StringBuilder();
                for (TextLayout.Run r : l.runs) lt.append(r.text);
                int at = lt.length() == 0 ? -1 : phon.text.indexOf(lt.toString(), phonPos);
                if (at >= 0) {
                    paintPhonetic(cv, l, at, lt.length(), x, y + rubyH + RUBY_DROP * phon.font.size + textDy / L.scale);
                    phonPos = at + lt.length();
                }
            }
            y += pitch(l) + rubyH + gap;
        }
    }

    /** 繰り返し文字のある表示形式 (会計など) を描く: 左の文字列、繰り返し、右の文字列。 */
    private void paintAccounting(PdfCanvas cv, TextLayout.Run base, ValueFormatter.Result vr, double x, double y0,
                                 double availW, double h, VerticalAlignment va) throws IOException {
        TextLayout.Run left = base.copy(vr.fillLeft), right = base.copy(vr.fillRight);
        float lw = TextLayout.width(left, ctx.fonts), rw = TextLayout.width(right, ctx.fonts);
        double asc = base.spec.ascent * base.size, desc = base.spec.descent * base.size;
        double baseline = switch (va) {
            case TOP -> y0 + padTop() + asc;
            case CENTER -> y0 + (h - asc - desc) / 2 + asc + centerDy();
            default -> y0 + h - padBottom() - desc;
        };
        paintPendingIcon(cv, baseline, base.size);
        if (suppressText) return;
        if ((vr.number || vr.logical) && lw + rw > availW + 0.01) {   // # で埋めるのは数値・真偽値だけ (文字列は埋めない)
            TextLayout.Run hash = base.copy("#");
            float hw = TextLayout.width(hash, ctx.fonts);
            int n = hw > 0 ? (int) Math.floor(availW / hw) : 0;
            cv.text("#".repeat(Math.max(0, n)), x, baseline, base.spec, base.size, base.color, 1, 0);
            return;
        }
        cv.text(left.text, x, baseline, base.spec, base.size, base.color, 1, 0);
        cv.text(right.text, x + availW - rw, baseline, base.spec, base.size, base.color, 1, 0);
        if (!" ".equals(vr.fillChar)) {
            TextLayout.Run f = base.copy(vr.fillChar);
            float fw = TextLayout.width(f, ctx.fonts);
            if (fw > 0) {
                int n = (int) Math.floor((availW - lw - rw) / fw);
                if (n > 0) cv.text(vr.fillChar.repeat(n), x + lw, baseline, base.spec, base.size, base.color, 1, 0);
            }
        }
    }

    /** 縦書き (文字を縦に積む) に使う Unicode の縦書き用字形。 */
    static final java.util.Map<Integer, Integer> VERTICAL_FORMS = java.util.Map.ofEntries(
            java.util.Map.entry(0xFF08, 0xFE35), java.util.Map.entry(0xFF09, 0xFE36),   // （ ）
            java.util.Map.entry(0x300C, 0xFE41), java.util.Map.entry(0x300D, 0xFE42),   // 「 」
            java.util.Map.entry(0x300E, 0xFE43), java.util.Map.entry(0x300F, 0xFE44),   // 『 』
            java.util.Map.entry(0x3010, 0xFE3B), java.util.Map.entry(0x3011, 0xFE3C),   // 【 】
            java.util.Map.entry(0x3014, 0xFE39), java.util.Map.entry(0x3015, 0xFE3A),   // 〔 〕
            java.util.Map.entry(0xFF3B, 0xFE47), java.util.Map.entry(0xFF3D, 0xFE48),   // ［ ］
            java.util.Map.entry(0xFF5B, 0xFE37), java.util.Map.entry(0xFF5D, 0xFE38),   // ｛ ｝
            java.util.Map.entry(0x3008, 0xFE3F), java.util.Map.entry(0x3009, 0xFE40),   // 〈 〉
            java.util.Map.entry(0x300A, 0xFE3D), java.util.Map.entry(0x300B, 0xFE3E),   // 《 》
            java.util.Map.entry(0x3001, 0xFE11), java.util.Map.entry(0x3002, 0xFE12),   // 、 。
            java.util.Map.entry(0x2026, 0xFE19), java.util.Map.entry(0x2025, 0xFE30));  // … ‥

    /** 縦書きで 90 度回す文字 (縦書き用字形がない): 長音、ダッシュ、波線など。 */
    static final String VERTICAL_ROTATE = "\u30FC\uFF0D\u2015\u2014\u2500\u301C\uFF5E\uFF1D\u2192\u2190";

    /** 句読点・リーダー (縦書き用字形がなくても回さない)。 */
    private static boolean isPunct(int cp) {
        return cp == 0x3001 || cp == 0x3002 || cp == 0x2026 || cp == 0x2025;
    }

    /** 縦書き (文字を上から下に積む) で描く。 */
    private void paintStacked(PdfCanvas cv, List<TextLayout.Run> runs, double x0, double y0, double w, double h,
                              HorizontalAlignment ha, VerticalAlignment va, double indentPt, boolean wrap) throws IOException {
        // 改行で列を分ける
        List<List<int[]>> colCps = new ArrayList<>();
        List<List<TextLayout.Run>> colOwners = new ArrayList<>();
        colCps.add(new ArrayList<>());
        colOwners.add(new ArrayList<>());
        for (TextLayout.Run r : runs) {
            for (int i = 0; i < r.text.length(); ) {
                int cp = r.text.codePointAt(i);
                i += Character.charCount(cp);
                if (cp == '\r') continue;
                if (cp == '\n') {
                    colCps.add(new ArrayList<>());
                    colOwners.add(new ArrayList<>());
                    continue;
                }
                colCps.get(colCps.size() - 1).add(new int[] {cp});
                colOwners.get(colOwners.size() - 1).add(r);
            }
        }
        // 折り返す設定なら、セルの高さに収まらない列をさらに分ける
        // 高さ 1 文字分のセルで「日 曜」と右から左に並ぶ)
        if (wrap) {
            double availH = h - padTop() - padBottom() + 1e-6;
            List<List<int[]>> cps2 = new ArrayList<>();
            List<List<TextLayout.Run>> own2 = new ArrayList<>();
            for (int k = 0; k < colCps.size(); k++) {
                List<int[]> cs = colCps.get(k);
                if (cs.isEmpty()) { cps2.add(cs); own2.add(colOwners.get(k)); continue; }
                // 空白の位置で折り返し、列の先頭の空白は詰める
                int i = 0;
                while (i < cs.size()) {
                    while (i < cs.size() && isStackSpace(cs.get(i)[0]) && i > 0) i++;
                    if (i >= cs.size()) break;
                    // 収まる文字数は文字ごとの字送りで数える
                    int e = i + 1;
                    double used = pitchOf(colOwners.get(k).get(i), runs);
                    while (e < cs.size()) {
                        double ph = pitchOf(colOwners.get(k).get(e), runs);
                        if (used + ph > availH) break;
                        used += ph;
                        e++;
                    }
                    if (e < cs.size() && !isStackSpace(cs.get(e)[0])) {
                        int sp = -1;
                        for (int q = e - 1; q > i; q--) if (isStackSpace(cs.get(q)[0])) { sp = q; break; }
                        if (sp > i) e = sp;
                    }
                    int ee = e;
                    while (ee > i && isStackSpace(cs.get(ee - 1)[0])) ee--;
                    if (ee > i) {
                        cps2.add(new ArrayList<>(cs.subList(i, ee)));
                        own2.add(new ArrayList<>(colOwners.get(k).subList(i, ee)));
                    }
                    i = e;
                }
            }
            colCps = cps2;
            colOwners = own2;
        }
        int n = colCps.size();
        if (n == 1) {
            paintStackedColumn(cv, colCps.get(0), colOwners.get(0), x0, y0, w, h, ha, va, indentPt);
            return;
        }
        // 列の間隔は折り返した行の行送りと同じ。列のまとまりを横位置に合わせて置く
        TextLayout.Run f = runs.get(0);
        double colPitch = yuFont(f) ? (f.spec.ascent + f.spec.descent) * f.size * ROT_PITCH : L.stackPitch(f.spec.requested, f.size);
        if (colPitch <= 0) colPitch = (f.spec.ascent + f.spec.descent) * f.size;
        double block = n * colPitch;
        double left = switch (ha) {
            case LEFT -> x0 + padX();
            case RIGHT -> x0 + w - padX() - block;
            default -> x0 + (w - block) / 2;
        };
        for (int k = 0; k < n; k++) {
            double cx0 = left + (n - 1 - k) * colPitch;
            paintStackedColumn(cv, colCps.get(k), colOwners.get(k), cx0, y0, colPitch, h, HorizontalAlignment.CENTER, va, indentPt);
        }
    }

    /**
     * セルに配置した画像 (値のメタデータ vm が画像を指すセル) を、縦横比を保ってセルに収め、セルの横位置・縦位置に合わせて描く。
     *
     * @return 画像を描いたか (画像でなければ false で、通常の値を描く)
     */
    static final double IMG_SHIFT = 0.36, IMG_SHRINK = 0.12;

    private boolean paintCellImage(PdfCanvas cv, XSSFCell cell, Align al, double x0, double y0, double w, double h) throws IOException {
        CellImages ci = CellImages.of(cell.getSheet().getWorkbook());
        long vm = cell.getCTCell().getVm();
        byte[] data = ci.image(vm);
        if (data == null) return false;
        int[] wh = Images.size(data);
        if (wh == null || wh[0] <= 0 || wh[1] <= 0) return false;
        // Excel の印刷は、セルの枠を左上へ 0.36pt ずらし 0.12pt 小さくした枠に収める (大きさ・配置によらない)
        double cx0 = x0, cy0 = y0, cw = w, ch = h;
        x0 -= IMG_SHIFT; y0 -= IMG_SHIFT; w = Math.max(0, w - IMG_SHRINK); h = Math.max(0, h - IMG_SHRINK);
        double k = Math.min(w / wh[0], h / wh[1]);
        double iw = wh[0] * k, ih = wh[1] * k;
        double x = switch (al.h()) {
            case LEFT -> x0;
            case RIGHT -> x0 + w - iw;
            case CENTER, CENTER_SELECTION, DISTRIBUTED, JUSTIFY, FILL -> x0 + (w - iw) / 2;
            default -> x0 + (w - iw) / 2;
        };
        double y = switch (al.v()) {
            case TOP -> y0;
            case CENTER, DISTRIBUTED, JUSTIFY -> y0 + (h - ih) / 2;
            default -> y0 + h - ih;
        };
        var img = Images.toPdf(ctx, "cellimage:" + ci.name(vm), data, ci.name(vm), iw, ih);
        if (img == null) return false;
        cv.save();
        try {
            cv.clipRect(cx0 - IMG_SHIFT, cy0 - IMG_SHIFT, cw, ch);
            cv.image(img, x, y, iw, ih, 1f);
        } finally {
            cv.restore();
        }
        return true;
    }

    /** 書式ごとの「セル内のチェックボックス」の判定のキャッシュ。 */
    private final Map<Integer, Boolean> checkboxXf = new HashMap<>();

    /**
     * セル内のチェックボックス (Microsoft 365 の機能。書式の extLst に xfpb:xfComplement を持ち、値は TRUE/FALSE) か。
     * Excel は値の代わりに四角 (チェックありならレ点) を描く。
     */
    private boolean isCheckboxCell(XSSFCellStyle st) {
        return checkboxXf.computeIfAbsent((int) st.getIndex(), k -> {
            var xf = st.getCoreXf();
            return xf.isSetExtLst() && xf.getExtLst().xmlText().contains("xfComplement");
        });
    }

    /** セル内のチェックボックスを描く: フォントの大きさの四角を、横位置 (標準は中央) と縦位置に合わせて置く。 */
    private void paintCellCheckbox(PdfCanvas cv, XSSFCell cell, XSSFFont font, Align al, double x0, double y0, double w, double h)
            throws IOException {
        double size = font.getFontHeight() / 20.0 * 0.9;
        HorizontalAlignment ha = al.h();
        double x = switch (ha) {
            case LEFT -> x0 + padX() + 1;
            case RIGHT -> x0 + w - padX() - 1 - size;
            default -> x0 + (w - size) / 2;
        };
        double y = switch (al.v()) {
            case TOP -> y0 + padTop() + 1;
            case CENTER, DISTRIBUTED, JUSTIFY -> y0 + (h - size) / 2;
            default -> y0 + h - padBottom() - 1 - size;
        };
        Color col = fontColor(font);
        cv.stroke(new java.awt.geom.Rectangle2D.Double(x, y, size, size), new PdfCanvas.Stroke((float) Math.max(0.7, size * 0.11)), col);
        if (cell.getBooleanCellValue()) {
            java.awt.geom.Path2D p = new java.awt.geom.Path2D.Double();
            p.moveTo(x + size * 0.2, y + size * 0.5);
            p.lineTo(x + size * 0.42, y + size * 0.75);
            p.lineTo(x + size * 0.8, y + size * 0.25);
            cv.stroke(p, new PdfCanvas.Stroke((float) Math.max(0.8, size * 0.12)), col);
        }
    }

    /** 縦書きの折り返しで空白として扱う文字か。 */
    private static boolean isStackSpace(int cp) {
        return cp == ' ' || cp == '\u3000';
    }

    /**
     * 縦書きの 1 文字の字送り (PrinterModel.stackPitchDev。游ゴシックも同じ)。
     */
    private double stackPitch(TextLayout.Run f) {
        double p = L.stackPitch(f.spec.requested, f.size);
        return p > 0 ? p : (f.spec.ascent + f.spec.descent) * f.size;
    }

    /** 回転した文字の行の間隔: 游ゴシックは (アセント + ディセント) × 1.07、ほかは折り返しの行送り。 */
    private double rotPitch(TextLayout.Line l) {
        double base = (l.ascent + l.descent) * ROT_PITCH;
        return base;
    }

    private static boolean yuFont(TextLayout.Run f) {
        String n = f.spec.requested == null ? "" : f.spec.requested.toLowerCase(java.util.Locale.ROOT);
        return n.startsWith("游ゴシック") || n.startsWith("yu gothic");
    }

    /** 縦書きで、ラン r の文字 1 つの字送り (文字ごとの大きさに従う。先頭に揃える切替では runs の先頭)。 */
    private double pitchOf(TextLayout.Run r, List<TextLayout.Run> runs) {
        return stackPitch(r);
    }

    /** 縦書きの 1 列を、横の範囲 x0〜x0+w に描く。 */
    private void paintStackedColumn(PdfCanvas cv, List<int[]> cps, List<TextLayout.Run> owners, double x0, double y0, double w, double h,
                                    HorizontalAlignment ha, VerticalAlignment va, double indentPt) throws IOException {
        if (cps.isEmpty()) return;
        double[] ph = new double[cps.size()];
        double total = 0;
        for (int i = 0; i < ph.length; i++) total += ph[i] = pitchOf(owners.get(i), owners);
        double top = switch (va) {
            case TOP -> y0 + padTop();
            // 中央揃えは STACK_CENTER_PX 下、結合セルはそこから MERGE_CENTER_PX 上 (書体・大きさによらず単独 5 画素、結合 1 画素)
            case CENTER -> y0 + (h - total) / 2
                    + (posx() ? (STACK_CENTER_PX - (mergedText ? MERGE_CENTER_PX : 0)) * 0.12 / Math.max(0.1, L.scale) : 0);
            case DISTRIBUTED, JUSTIFY -> y0 + (h - total) / 2;
            default -> y0 + h - padBottom() - total;
        };
        double gap = 0;
        if ((va == VerticalAlignment.DISTRIBUTED || va == VerticalAlignment.JUSTIFY) && cps.size() > 1 && total < h - padTop() - padBottom()) {
            // 縦位置が均等割り付けのときは、字下げの幅を上下の余白にする
            double ind = va == VerticalAlignment.DISTRIBUTED ? Math.min(indentPt, Math.max(0, (h - total) / 2)) : 0;
            top = y0 + padTop() + ind;
            gap = (h - padTop() - padBottom() - 2 * ind - total) / (cps.size() - 1);
        }
        double cur = top;
        for (int i = 0; i < cps.size(); i++) {
            TextLayout.Run r = owners.get(i);
            double at = cur, lh = ph[i];
            cur += lh + gap;
            int cp = cps.get(i)[0];
            // 縦書きは縦書き用字形 (（ → ︵ など) を使うか、グリフを 90 度回す (ー、～)
            Integer vf = VERTICAL_FORMS.get(cp);
            boolean turn = VERTICAL_ROTATE.indexOf(cp) >= 0;
            if (vf != null) {
                // 同じフォントの縦書き用字形。なければ横書きのグリフを回す (書体を保つ)
                if (r.spec.faceFor(cp, ctx.fonts).hasGlyph(vf)) cp = vf;
                else turn = !isPunct(cp);
            }
            String ch = new String(Character.toChars(cp));
            float cw = r.spec.charWidth(cp, r.size, ctx.fonts);
            if (turn) {
                double asc = r.spec.ascent * r.size, desc = r.spec.descent * r.size;
                double cx = switch (ha) {
                    case LEFT -> x0 + padX() + (asc + desc) / 2;
                    case RIGHT -> x0 + w - padX() - (asc + desc) / 2;
                    default -> x0 + w / 2;
                };
                double cy = at + lh / 2;
                cv.glyph(ch, cx + (desc - asc) / 2, cy - cw / 2, Math.PI / 2, r.spec, r.size, r.color);
                continue;
            }
            double x = switch (ha) {
                case LEFT -> x0 + padX();
                case RIGHT -> x0 + w - padX() - cw;
                default -> x0 + (w - cw) / 2;
            };
            double baseline = at + r.spec.ascent * r.size;
            cv.text(ch, x, baseline, r.spec, r.size, r.color, 1, 0);
        }
    }

    /** 文字を角度を付けて描く。 */
    /** 折り返す回転した文字の現在のセル (paintCellText で決める)。 */
    private boolean rotWrap;

    private void paintRotated(PdfCanvas cv, List<TextLayout.Run> runs, double x0, double y0, double w, double h,
                              HorizontalAlignment ha, VerticalAlignment va, double angleDeg, double padL, double padR)
            throws IOException {
        // 回転した文字は、折り返しの設定によらず改行で行を分ける (Excel の印刷。回転しないセルは折り返しなしなら改行を無視する)
        List<TextLayout.Line> lines = TextLayout.layout(runs, 0, false, true, ctx.fonts);
        if (lines.isEmpty()) return;
        if (rotWrap) {
            // 折り返す回転した文字: 1 行の縦の広がり (長さ × sin + 行の高さ × cos) がセルの高さ − 上下の余白に収まる長さで折り返す
            // (3 つの角度の折り返し位置が合う)
            double ra = Math.toRadians(Math.abs(angleDeg));
            double lh = lines.get(0).ascent + lines.get(0).descent;
            double len = (h - padTop() - padBottom() - lh * Math.cos(ra)) / Math.max(0.05, Math.sin(ra));
            if (len > 1) lines = TextLayout.layout(runs, (float) len, true, true, ctx.fonts);
        }
        TextLayout.Line l = lines.get(0);
        double tw = 0, th = 0;
        for (int i = 0; i < lines.size(); i++) {
            TextLayout.Line li = lines.get(i);
            if (li.ascent + li.descent == 0) { li.ascent = l.ascent; li.descent = l.descent; }
            tw = Math.max(tw, li.width);
            th += i == lines.size() - 1 ? li.ascent + li.descent : pitch(li);
        }
        double a = Math.toRadians(angleDeg);
        double cos = Math.abs(Math.cos(a)), sin = Math.abs(Math.sin(a));
        if (sin >= 0.999) {
            paintRot90(cv, lines, x0, y0, w, h, ha, va, angleDeg > 0);
            return;
        }
        if (slantBox && sin < 0.999 && sin > 0.01) {
            paintSlanted(cv, lines, x0, y0, w, h, ha, va, angleDeg);
            return;
        }
        if (lines.size() > 1 && sin > 0.2) {
            paintRotatedLines(cv, lines, x0, y0, w, h, ha, va, a, padL, padR);
            return;
        }
        double bw = tw * cos + th * sin, bh = tw * sin + th * cos;
        double cx = switch (ha) {
            case LEFT -> x0 + padL + bw / 2;
            case RIGHT -> x0 + w - padR - bw / 2;
            default -> x0 + w / 2;
        };
        double cy = switch (va) {
            case TOP -> y0 + padTop() + bh / 2;
            case CENTER, JUSTIFY, DISTRIBUTED -> y0 + h / 2;
            default -> y0 + h - padBottom() - bh / 2;
        };
        cv.save();
        try {
            java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(cx, cy);
            at.rotate(-a); // Excel の正の角度は反時計回り
            cv.transform(at);
            double yy = -th / 2;
            for (TextLayout.Line li : lines) {
                // 行の揃えは、まとまりの中で横の配置に合わせる
                double lx = switch (ha) {
                    case LEFT -> -tw / 2;
                    case RIGHT -> tw / 2 - li.width;
                    default -> -li.width / 2;
                };
                TextLayout.drawLine(cv, li, lx, yy + li.ascent, 0, 0);
                yy += pitch(li);
            }
        } finally {
            cv.restore();
        }
    }

    /**
     * 複数行の回転した文字: 各行の書き始めの点を同じ高さに揃え、行の間隔 ÷ sin(角度) ずつ横にずらす
     * (行の間隔は文字のアセント + ディセントの 1.07 倍。游ゴシック 11pt で 18.48pt)。全体の外枠をセルの配置に合わせて置く。
     */
    static final double ROT_PITCH = 1.07, ROT_ASC_EXTRA = 1.11;
    /** 游ゴシック 11pt のアセント + ディセント (pt)。回転の配置の数値はこの大きさのもの。 */
    static final double YU11_AD = 18.48 / 1.07;
    static final double R90_TOP = 0.56, R90_BOT = 0.63;
    static final double M90_TOP = 2.58, M90_BOT = 0;

    /**
     * ±90 度の文字 (罫線の有無によらない。数値は游ゴシック 11pt のもので文字の大きさに比例)。
     * 行は +90 度は右へ、−90 度は左へ (アセント + ディセント) × 1.07 ずつ並べる。横の配置: +90 度は左で 1 行目の基線をセルの左 + 13.54、
     * 右で最後の行をセルの右 − 9.1。−90 度は左で最後の行をセルの左 + 5.65、右で 1 行目をセルの右 − 16.74。中央はその間。
     * 縦の配置は行ごとに当てる: +90 度は下で書き始めを下 − 0.63、上で書き終わりを上 + 0.56。−90 度は上で書き始めを上 + 2.58、
     * 下で書き終わりを下。
     */
    private void paintRot90(PdfCanvas cv, List<TextLayout.Line> lines, double x0, double y0, double w, double h,
                            HorizontalAlignment ha, VerticalAlignment va, boolean up) throws IOException {
        int n = lines.size();
        TextLayout.Line l0 = lines.get(0);
        double k = (l0.ascent + l0.descent) / YU11_AD;
        double pitch = rotPitch(l0);
        double x1 = x0 + w, yB = y0 + h;
        boolean left = ha == HorizontalAlignment.LEFT || ha == HorizontalAlignment.FILL, right = ha == HorizontalAlignment.RIGHT;
        if (ha == HorizontalAlignment.GENERAL) { left = !up; right = up; }
        double first;
        if (up) {
            double xl = x0 + 13.54 * k, xr = x1 - 9.1 * k - (n - 1) * pitch;
            first = left ? xl : right ? xr : (xl + xr) / 2;
        } else {
            double xr = x1 - 16.74 * k, xl = x0 + 5.65 * k + (n - 1) * pitch;
            first = right ? xr : left ? xl : (xl + xr) / 2;
        }
        boolean top = va == VerticalAlignment.TOP, bottom = va == VerticalAlignment.BOTTOM;
        for (int i = 0; i < n; i++) {
            TextLayout.Line li = lines.get(i);
            double x = first + (up ? i : -i) * pitch, y;
            if (up) {
                double yb = yB - R90_BOT * k, yt = y0 + R90_TOP * k + li.width;
                y = bottom ? yb : top ? yt : (yb + yt) / 2;
            } else {
                double yt = y0 + M90_TOP * k, yb = yB - M90_BOT * k - li.width;
                y = top ? yt : bottom ? yb : (yb + yt) / 2;
            }
            cv.save();
            try {
                java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(x, y);
                at.rotate(up ? -Math.PI / 2 : Math.PI / 2);
                cv.transform(at);
                TextLayout.drawLine(cv, li, 0, 0, 0, 0);
            } finally {
                cv.restore();
            }
        }
    }

    /** 描いている斜めの文字のセルに罫線がある (Excel は罫線のあるセルだけ平行四辺形として文字を置く)。 */
    private boolean slantBox;

    /**
     * 斜めの文字 (0 < |角度| < 90) を Excel の印刷と同じに置く (数値は游ゴシック 11pt のもので、文字の大きさに比例させる)。
     * セルを、下の辺 (負の角度は上の辺) と文字の向きの辺で張る平行四辺形とみなし、横の位置 a (水平) と文字の向きの位置 b で表す。
     * <ul>
     *   <li>行は a の向きに行の間隔 ÷ sin ずつ並べる。横の配置: 左は a = −2.8 + 20.14 / sin、右は a = 幅 − 6.67 + 1.69 / sin
     *       (最後の行)、中央はその間。負の角度は左右を入れ替え、δ = 1.17 − 4.15 sin を足す。</li>
     *   <li>縦の配置は b: 下は 1.14 − 0.834 cot、上は 長さ − 最も長い行の幅 − (19.37 cot + 2.32)、中央はその間 (負の角度は上下が逆)。
     *       各行はその中で横の配置に合わせて揃える。</li>
     * </ul>
     */
    private double paintSlanted(PdfCanvas cv, List<TextLayout.Line> lines, double x0, double y0, double w, double h,
                                HorizontalAlignment ha, VerticalAlignment va, double angleDeg) throws IOException {
        boolean neg = angleDeg < 0;
        double t = Math.toRadians(Math.abs(angleDeg)), sn = Math.sin(t), cs = Math.cos(t), cot = cs / sn;
        int n = lines.size();
        double wmax = 0;
        for (TextLayout.Line li : lines)
            wmax = Math.max(wmax, li.width);
        TextLayout.Line l0 = lines.get(0);
        double k = (l0.ascent + l0.descent) / YU11_AD;
        double pitch = rotPitch(l0), step = pitch / sn;
        double len = h / sn, ext = (19.37 * cot + 2.32) * k, bLow = (1.14 - 0.834 * cot) * k;
        // 横 (a): 正の角度の左・右
        // 45 度を超えると左へ (60 度で 1.28pt)
        double steep = Math.max(0, Math.abs(angleDeg) - 45) / 15 * 1.28 * k;
        double aL = -2.8 * k + 20.14 * k / sn - steep, aR = w - 6.67 * k + 1.69 * k / sn - (n - 1) * step - steep;
        boolean left = ha == HorizontalAlignment.LEFT || ha == HorizontalAlignment.GENERAL || ha == HorizontalAlignment.FILL;
        boolean right = ha == HorizontalAlignment.RIGHT;
        double a1;
        if (!neg) a1 = left ? aL : right ? aR : (aL + aR) / 2;
        else {
            double d = (1.17 - 4.15 * sn) * k;
            // 負の角度は左右を入れ替え、行は左へ並べる (1 行目が右)
            double aR1 = w - 6.67 * k + 1.69 * k / sn;
            double nl = w - aR1 + d + (n - 1) * step, nr = w - aL + d;
            a1 = left ? nl : right ? nr : (nl + nr) / 2;
            if (!right) a1 += 0.322 * (n - 1) * step; // 複数行の左・中央は右へ
        }
        // 縦 (b): 文字の向きに沿った、まとまりの始まり
        boolean top = va == VerticalAlignment.TOP, bottom = va == VerticalAlignment.BOTTOM;
        double bStart = bLow, bEnd = len - wmax - ext; // 正の角度の下・上
        if (neg) { double tmp = ext; bStart = tmp; bEnd = len - wmax - bLow; } // 負の角度は上から下へ: 上・下
        double bBlk = !neg ? (bottom ? bStart : top ? bEnd : (bStart + bEnd) / 2)
                : (top ? bStart : bottom ? bEnd : (bStart + bEnd) / 2);
        double yB = y0 + h, xTop = x0 - h * cot, maxX = -Double.MAX_VALUE;
        for (int i = 0; i < n; i++) {
            TextLayout.Line li = lines.get(i);
            double b = bBlk + (left ? 0 : right ? wmax - li.width : (wmax - li.width) / 2);
            if (sn >= 0.999) b = top ? len - li.width - ext : bottom ? bLow : (len - li.width - ext + bLow) / 2; // 90 度は行ごとに縦の配置
            double a = a1 + (neg ? -i : i) * step;
            double ox = neg ? xTop + a + b * cs : x0 + a + b * cs;
            double oy = neg ? y0 + b * sn : yB - b * sn;
            maxX = Math.max(maxX, ox + li.width * cs + li.descent * sn);
            if (cv == null) continue; // 測るだけ
            cv.save();
            try {
                java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(ox, oy);
                at.rotate(-Math.toRadians(angleDeg));
                cv.transform(at);
                TextLayout.drawLine(cv, li, 0, 0, 0, 0);
            } finally {
                cv.restore();
            }
        }
        return maxX;
    }

    private void paintRotatedLines(PdfCanvas cv, List<TextLayout.Line> lines, double x0, double y0, double w, double h,
                                   HorizontalAlignment ha, VerticalAlignment va, double a, double padL, double padR) throws IOException {
        double sa = Math.sin(a), ca = Math.cos(a);
        double ux = ca, uy = -sa, vx = -sa, vy = -ca; // 文字の向き、上の向き (用紙の座標、y は下向き)
        int n = lines.size();
        double[] ox = new double[n];
        double minX = Double.MAX_VALUE, maxX = -Double.MAX_VALUE, minY = Double.MAX_VALUE, maxY = -Double.MAX_VALUE, acc = 0;
        for (int i = 0; i < n; i++) {
            TextLayout.Line li = lines.get(i);
            ox[i] = acc;
            acc += rotPitch(li) / Math.abs(sa) * Math.signum(sa);
            for (double k : new double[] {0, li.width}) for (double m : new double[] {li.ascent, -li.descent}) {
                double px = ox[i] + k * ux + m * vx, py = k * uy + m * vy;
                minX = Math.min(minX, px); maxX = Math.max(maxX, px); minY = Math.min(minY, py); maxY = Math.max(maxY, py);
            }
        }
        double bw = maxX - minX, bh = maxY - minY;
        double left = switch (ha) {
            case LEFT -> x0 + padL;
            case RIGHT -> x0 + w - padR - bw;
            default -> x0 + (w - bw) / 2;
        };
        double top = switch (va) {
            case TOP -> y0 + padTop();
            case CENTER, JUSTIFY, DISTRIBUTED -> y0 + (h - bh) / 2;
            default -> y0 + h - padBottom() - bh;
        };
        // Excel は行の上側に文字の大きさ × 1.11 の高さを見込んで上下に揃える (30・45 度で 5.2・4.4pt 下)
        double dx = left - minX, dy = top - minY + (va == VerticalAlignment.CENTER ? 0.5 : va == VerticalAlignment.TOP ? 1 : 0)
                * (lines.get(0).ascent + lines.get(0).descent) / YU11_AD * 11 * ROT_ASC_EXTRA * Math.abs(ca);
        for (int i = 0; i < n; i++) {
            cv.save();
            try {
                java.awt.geom.AffineTransform at = java.awt.geom.AffineTransform.getTranslateInstance(ox[i] + dx, dy);
                at.rotate(-a);
                cv.transform(at);
                TextLayout.drawLine(cv, lines.get(i), 0, 0, 0, 0);
            } finally {
                cv.restore();
            }
        }
    }

    /**
     * 文字のベースラインを用紙上で下げる量 (pt)。Excel 2013 以前で保存したブックは、どの環境で印刷しても
     * 文字が 0.5pt 下に出る。
     */
    private double textDy() {
        int av = ctx.printer.appVersion;
        // 古い Excel で出力した PDF の目印 (--author-env のときだけ。今の Excel は AppVersion を見ない)
        return ctx.printer.authorEnv && av > 0 && av <= 15 ? LEGACY_TEXT_DY : 0;
    }
}
