package xlsx2pdf.render;

import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetFormatPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

/**
 * シートの列幅・行高 (シートの pt)、結合セル、使用範囲。
 *
 * <p>列幅・行高は、画面のピクセルを Excel の式で求め、{@link PrinterModel} の係数で印刷時のデバイスピクセルに
 * 換算してから pt にする。印刷倍率を変えるたびに {@link #setScale} で計算し直す。
 */
public final class SheetLayout {
    /** 96 dpi の 1 ピクセルの pt。 */
    public static final double PX = 0.75;

    /** 対象のシート。 */
    public final XSSFSheet sheet;
    /** 画面 (96 dpi) の最大数字幅 (px)。 */
    public final int mdw;
    /** 列幅・行高 (シートの pt、現在の倍率)。 */
    public final double[] colW, rowH;
    /** 画面の列幅 (pt。Excel の Columns.Width と同じ。値の書き出し用)。 */
    private final double[] colScr;
    /** 倍率で丸める前 (100%) の列幅 (シートの pt。「ページに合わせる」の倍率の上限に使う)。 */
    private final double[] colW100;
    /** 列・行の境界の位置 (累積和。要素数は列数・行数 + 1)。 */
    private final double[] colX, rowY;
    /** 配列の範囲 (この列・行より先は既定の大きさで外挿する)。 */
    public final int maxCol, maxRow;
    /** 既定の行高 (pt、現在の倍率)。 */
    public double defaultRowH;
    /** 結合範囲に含まれるすべてのセル (行 × 16384 + 列) → 結合範囲。 */
    private final Map<Long, CellRangeAddress> merged = new HashMap<>();
    /** 結合範囲の一覧。 */
    public final List<CellRangeAddress> mergedList;

    private final PrinterModel model;
    private final int hdpi, vdpi;
    /** 保存された列幅 (文字数。-1 は非表示)。 */
    private final double[] colChars;
    /** 行高 (画面ピクセル。整数ピクセルでない高さは端数あり。0 は非表示)。 */
    private final double[] rowPx;
    /** 既定の行高 (画面ピクセル)。 */
    private double defaultRowPx;
    /** 既定の列幅 (文字数)。 */
    private double defaultColChars;
    /** 既定の行高がユーザー指定 (customHeight) か。 */
    private final boolean defaultCustomHeight;
    /** sheetFormatPr の x14ac:dyDescent (なければ -1)。 */
    private double dyDescent = -1;
    /** 高さが保存されているが固定していない行の、行の書式と空でなく結合されていないセルの "フォント|サイズ" (なければ null)。 */
    private final String[][] rowFonts;

    /** 画面 1 pt が印刷で何 pt になるか (横・縦。描画のオフセットに使う)。 */
    public double xScale = 1, yScale = 1;
    /** 現在の印刷倍率 (1 = 100%)。 */
    public double scale = 1;
    /** 100% のレイアウトをそのまま縮小する (フォントを倍率で実体化し直さない。「ページに合わせる」の一部のシート、{@link Paginator#paginate})。 */
    public boolean linear;

    /** フォントを実体化する倍率 (linear なら 1)。 */
    public double realScale() { return linear ? 1.0 : scale; }
    public String dbgInfo() { PrinterModel.Factors f = model.factors(realScale(), hdpi, vdpi); return "authorDpi=" + authorDpi + " inferred=" + inferredDpi + " rowRatio=" + f.rowRatio + " rowPx0=" + rowPx[0] + " defCustom=" + defaultCustomHeight + " defRowPx=" + defaultRowPx + " linear=" + linear + " scale=" + scale; }

    /** 作成者 (出力した PC) の画面 DPI。保存された列幅から推定する (96 = 標準)。 */
    private int authorDpi = 96;
    /** 列幅から推定した作成者 DPI (画面 DPI の指定に関係なく)。 */
    private int inferredDpi = 96;
    /** 指定された出力環境の画面 DPI (0 = 推定する)。 */
    private final int screenDpi;
    /** authorDpi での最大数字幅と標準行高 (px)。 */
    private int authorMdw, authorRow;

    /** 高さが保存されているが手動で固定していない行 (Excel は開くときに自動調整し直す)。 */
    private boolean[] rowHtOnly;
    /**
     * 作成者の既定の行高が変換する環境の標準行と違うシートで、customHeight のない 1 行分の行を、変換する環境の画面で自動調整し直す
     * フォントは空のセル・1 行の結合セルも含む全セルと行の書式。
     * 保存値がその ±25% にあるときだけ (それより高い行は複数行)。
     */
    /**
     * customHeight のない行に 2 行以上に折り返すセルがあれば、Excel と同じく変換する環境の画面で行数を数え直し、行数 × フォントの画面の標準行に
     * する。画面の文字は ＭＳ 明朝・ＭＳ ゴシック (固定幅) なら「画面の ppem + 1」、ほかは ppem の大きさで測る (z_kekkaichiran 系 3 件 2464 行で 2451 行一致)。
     * 単独の CR では改行しない、選択範囲内で中央のセルは数えない、上限 409.5pt。
     * 未解決: Excel が測り直さず保存値のままの行もある。
     */
    static final int REWRAP = 5;
    static final double HTONLY_REFIT = 0.25;
    /** customHeight のない行に掛ける比 (既定の行高を標準行にしたときの比。それ以外は 1)。 */
    private double htOnlyRatio = 1;
    /** 保存された既定の行高 (手で決めていないもの) が、変換する環境の画面の標準行と違う。 */
    private boolean defaultDiffers;
    /** COLSTD: 列の書式で決まる、変換する環境の画面の標準行 (px。使わないときは 0)。 */
    private int colStdPx;
    /** 保存された既定の行高 (pt)。 */
    private double dhRaw;
    /** 行が customHeight か。 */
    private boolean[] rowCustom;
    /** 高さが保存されている行か。 */
    private boolean[] rowHasHt;

    /**
     * シートのレイアウトを作る (倍率 100% で計算する)。
     *
     * @param mdw      画面 (96 dpi) の最大数字幅 (px)
     * @param model    印刷の換算モデル
     * @param hdpi     横の解像度
     * @param vdpi     縦の解像度
     * @param needRows 計算が必要な行数
     * @param needCols 計算が必要な列数
     */
    public SheetLayout(XSSFSheet sheet, int mdw, PrinterModel model, int hdpi, int vdpi, int needRows, int needCols) {
        this(sheet, mdw, model, hdpi, vdpi, needRows, needCols, 0);
    }

    /**
     * 出力環境の画面 DPI を指定してシートのレイアウトを作る。
     *
     * @param screenDpi 出力した PC の画面 DPI (96 = 100% 表示、120 = 125%、144 = 150%、192 = 200%)。0 なら列幅から推定する
     */
    /** 数式の表示 (sheetView の showFormulas) で保存されたシートか。 */
    public final boolean showFormulas;

    /** 最初のシートビューが数式の表示か。 */
    public static boolean showFormulas(XSSFSheet sheet) {
        var svs = sheet.getCTWorksheet().getSheetViews();
        if (svs == null || svs.sizeOfSheetViewArray() == 0) return false;
        var v = svs.getSheetViewArray(0);
        return v.isSetShowFormulas() && v.getShowFormulas();
    }

    public SheetLayout(XSSFSheet sheet, int mdw, PrinterModel model, int hdpi, int vdpi, int needRows, int needCols, int screenDpi) {
        // 画面 DPI の指定がなければ、今の Excel・表示倍率 100% の PC (96 dpi) とする。--author-env なら列幅などから推定する
        this.screenDpi = screenDpi > 0 ? screenDpi : model.authorEnv ? 0 : 96;
        this.sheet = sheet;
        this.mdw = mdw;
        this.model = model;
        this.hdpi = hdpi;
        this.vdpi = vdpi;
        this.maxCol = Math.min(16384, Math.max(needCols, 1) + 1);
        this.maxRow = Math.min(1048576, Math.max(needRows, 1) + 1);
        CTWorksheet ws = sheet.getCTWorksheet();
        this.showFormulas = showFormulas(sheet);
        CTSheetFormatPr fp = ws.getSheetFormatPr();
        if (fp != null && fp.isSetDefaultColWidth()) {
            defaultColChars = fp.getDefaultColWidth();
        } else {
            long base = fp != null && fp.isSetBaseColWidth() ? fp.getBaseColWidth() : 8;
            // Excel の既定幅 = baseColWidth × mdw + 余白 5 px を、8 の倍数 (画面 px) に切り上げたもの
            double px = Math.ceil((base * mdw + 5) / 8.0) * 8;
            defaultColChars = px / mdw;
        }
        colChars = new double[maxCol];
        Arrays.fill(colChars, defaultColChars);
        boolean[] widthSet = new boolean[maxCol];
        for (CTCols cols : ws.getColsArray()) {
            for (CTCol c : cols.getColArray()) {
                int min = (int) c.getMin() - 1, max = (int) Math.min(c.getMax(), maxCol) - 1;
                for (int i = min; i <= max && i < maxCol; i++) {
                    if (c.isSetHidden() && c.getHidden()) colChars[i] = -1;
                    else if (c.isSetWidth()) { colChars[i] = c.getWidth(); widthSet[i] = true; }
                }
            }
        }
        if (ws.sizeOfColsArray() > 0 && this.screenDpi == 96 && !model.authorEnv) numWiden(widthSet);
        double dh = fp != null && fp.getDefaultRowHeight() > 0 ? fp.getDefaultRowHeight() : sheet.getDefaultRowHeightInPoints();
        defaultRowPx = toPx(dh);
        dhRaw = dh;
        defaultDiffers = !(fp != null && fp.isSetCustomHeight() && fp.getCustomHeight()) && Math.abs(defaultRowPx - model.screenRow) > 0.01;
        boolean zeroHeight = fp != null && fp.isSetZeroHeight() && fp.getZeroHeight();
        rowPx = new double[maxRow];
        Arrays.fill(rowPx, zeroHeight ? 0 : defaultRowPx);
        boolean defaultCustom = fp != null && fp.isSetCustomHeight() && fp.getCustomHeight();
        defaultCustomHeight = defaultCustom;
        if (fp != null) {
            org.w3c.dom.Node dn = fp.getDomNode().getAttributes().getNamedItemNS("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "dyDescent");
            if (dn != null) try { dyDescent = Double.parseDouble(dn.getNodeValue()); } catch (NumberFormatException ignored) { dyDescent = -1; }
        }
        detectAuthorDpi(ws, fp);
        if (!defaultCustom && authorDpi == 96
                && (authorDefaultValue(dh, dyDescent)
                    || model.latinRow > 0 && model.latinRow != model.screenRow && Math.abs(defaultRowPx - model.latinRow) < 0.01)) {
            // 手で決めていない既定の行高は使わず、標準フォントの画面の標準行にする
            double c = model.screenRow;
            for (int i = 0; i < rowPx.length; i++) if (rowPx[i] == defaultRowPx) rowPx[i] = c;
            htOnlyRatio = c / defaultRowPx;
            defaultRowPx = c;
        } else if (model.authorEnv && !defaultCustom && authorDpi == 96 && !authorDefault(dh, dyDescent)) {
            // 既定の行高 (手で決めていないもの) は画面の画素に切り上げて使う
            double c = Math.ceil(defaultRowPx - 1e-6);
            if (c != defaultRowPx) {
                for (int i = 0; i < rowPx.length; i++) if (rowPx[i] == defaultRowPx) rowPx[i] = c;
                defaultRowPx = c;
            }
        }
        colStdPx = !model.authorEnv && this.screenDpi == 96 && !defaultCustom ? colStdPx(ws) : 0;
        if (colStdPx > 0 && colStdPx != defaultRowPx) {
            for (int i = 0; i < rowPx.length; i++) if (rowPx[i] == defaultRowPx) rowPx[i] = colStdPx;
            defaultRowPx = colStdPx;
        }
        if (authorDpi != 96 && !(fp != null && fp.isSetDefaultColWidth())) {
            // 既定の列幅は作成者の画面で決まる (144 dpi の ＭＳ 明朝 11pt: 8 × 11 + 5 → 96 px = 8.73 文字)
            long base = fp != null && fp.isSetBaseColWidth() ? fp.getBaseColWidth() : 8;
            double adef = Math.ceil((base * authorMdw + 5) / 8.0) * 8 / authorMdw;
            for (int i = 0; i < maxCol; i++) if (!widthSet[i] && colChars[i] >= 0) colChars[i] = adef;
            defaultColChars = adef;
        }
        rowFonts = new String[maxRow][];
        mergedList = sheet.getMergedRegions();
        for (CellRangeAddress a : mergedList) {
            if (a.getNumberOfCells() > 2_000_000) continue;
            for (int r = a.getFirstRow(); r <= a.getLastRow(); r++)
                for (int c = a.getFirstColumn(); c <= a.getLastColumn(); c++) merged.put(key(r, c), a);
        }
        if (VALUE_FIT && this.screenDpi == 96 && !model.authorEnv) valueFit(ws);   // 結合セルは除くので、結合の表の後
        rowHtOnly = new boolean[maxRow];
        rowHasHt = new boolean[maxRow];
        rowCustom = new boolean[maxRow];
        for (int r = 0; r < maxRow; r++) {
            XSSFRow row = sheet.getRow(r);
            if (row == null) continue;
            boolean custom = row.getCTRow().isSetCustomHeight() && row.getCTRow().getCustomHeight();
            rowCustom[r] = custom;
            if (row.getZeroHeight()) { rowPx[r] = 0; continue; }
            else if (row.getCTRow().isSetHt() && !(model.nonExcel && !custom)) {
                rowPx[r] = toPx(row.getCTRow().getHt());
                rowHasHt[r] = true;
                if (!custom) rowPx[r] *= htOnlyRatio;
                if (!custom && !model.authorEnv && this.screenDpi == 96) rowPx[r] = Math.ceil(rowPx[r] - 1e-6);
                int rw = !custom && !model.authorEnv && this.screenDpi == 96 ? rewrapPx(row) : 0;
                if (rw > 0) rowPx[r] = rw;
                else if (!custom && (htOnlyRatio != 1 || colStdPx > 0 && colStdPx != toPx(dhRaw))
                        && !model.authorEnv && this.screenDpi == 96) {
                    int fit = refitPx(row);
                    if (colStdPx > 0 && fit > 0 && !(ROWFMT_NOSTD && row.isFormatted() && row.getRowStyle() != null)) fit = Math.max(fit, colStdPx);
                    double stored = toPx(row.getCTRow().getHt());
                    if (fit > 0 && Math.abs(stored - fit) <= HTONLY_REFIT * fit) rowPx[r] = fit;
                }
                if (rw <= 0 && !custom && !model.authorEnv && this.screenDpi == 96) {
                    skipMerged = AUTO_SKIP_MERGE;
                    int fit = autoRowPx(row);
                    skipMerged = false;
                    // 保存値が書体の標準行の最大以上なら、Excel の自動調整の結果 (基準線の和まで上げない。exp123: 游ゴシックのブックの
                    // BIZ UDPゴシック 18pt の行は 28px のまま)。低いときだけ基準線の和まで上げる (exp30 sei9)
                    double st = toPx(row.getCTRow().getHt());
                    if (fit > 0 && st < fit + 0.5 && !(PLAIN_KEEP && st >= lastPlainFit - 0.01)) rowPx[r] = fit;
                    // 折り返しのセルが全部 1 行に収まるなら、保存値より低くても測り直す
                    else if (fit > 0 && (hasWrapText(row) || SEL_WRAP_KEEP && hasSelWrap(row) && st < fit * 1.5)) rowPx[r] = fit;
                }
            }
            else {
                rowPx[r] = defaultRowPx;
                // 既定の行高が固定 (customHeight) でも、行の要素があって ht がなければ自動調整する
                if (defaultCustom && row.getPhysicalNumberOfCells() > 0 && !model.authorEnv && this.screenDpi == 96) {
                    int fit = autoRowPx(row);
                    if (fit > 0) rowPx[r] = fit;
                }
                if (colStdPx > 0 && row.getPhysicalNumberOfCells() > 0) {
                    // 高さの保存されていない行は、列の書式で決まる標準行とセルのフォントの大きい方
                    int rw = 0;
                    int fit = refitPx(row, false, true);
                    if (fit > 0) rowPx[r] = Math.max(colStdPx, Math.max(rw, fit));
                    // 行のふりがなの表示 (row の ph="1"): Excel は開くときにふりがなの 1 行分を足す (夜間 n3716: 30 + 13 = 43px = 32.25pt)
                    if (fit > 0 && ROW_PH && row.getCTRow().isSetPh() && row.getCTRow().getPh()) rowPx[r] += sheetRubyPx();
                } else { }
                // 高さの保存されていない行で、ふりがなを表示するセル (結合セルを除く) があるもの: Excel は開くときにふりがなの分を
                // 足して自動調整する。既定より高くなるときだけ使う。折り返しや大きな文字だけの行は既定の高さのまま (実ファイル)
                if (!defaultCustom && !colPhRuby && hasRuby(row) && !model.authorEnv && this.screenDpi == 96) {
                    int fit = rewrapPx(row);
                    if (fit <= 0) fit = autoRowPx(row);
                    if (fit > rowPx[r] + 0.5) rowPx[r] = fit;
                }
                if (NOHT_REWRAP > 0 && !defaultCustom && !model.authorEnv && this.screenDpi == 96 && !row.getCTRow().isSetHt()) {
                    // 高さの保存されていない行でも、折り返して 2 行以上になるセルがあれば Excel は開くときに行数分の高さにする
                    // (夜間 n1299: 標準フォントの数字幅が作成者 9 → この PC 6 で列が狭くなり、8pt の 65 文字が 2 行になる)
                    int rw = rewrapPx(row);
                    if (rw > rowPx[r] + 0.5) rowPx[r] = rw;
                }
                if (model.nonExcel && row.getCTRow().isSetHt() && !custom) {
                    // Excel 以外で保存したブックの手で決めていない高さは、Excel が開くときに内容で測り直す (折り返しを含む。exp111)
                    int rw = rewrapPx(row);
                    int fit = rw > 0 ? rw : autoRowPx(row);
                    if (fit > 0) rowPx[r] = fit;
                }
            }
            // 高さが保存された行はその高さを (換算して) 使う。ただし手動で固定していない行は、印刷する PC の画面 DPI が作成者と違えば
            // 自動調整し直す ({@link #setScale})
            rowHtOnly[r] = !custom && row.getCTRow().isSetHt() && !defaultCustom;
            if (rowHtOnly[r]) {
                rowFonts[r] = fontsOf(row);
                // Excel の自動調整は、文字のあるセルのフォントと、行の書式 (customFormat。なければ標準フォント) のフォントの最大
                String base = model.defaultFont() + "|" + model.defaultSize();
                if (row.isFormatted() && row.getRowStyle() != null) {
                    org.apache.poi.xssf.usermodel.XSSFFont rf = ((org.apache.poi.xssf.usermodel.XSSFCellStyle) row.getRowStyle()).getFont();
                    base = fkey(rf);
                }
                java.util.List<String> fl = new java.util.ArrayList<>(java.util.List.of(base));
                if (rowFonts[r] != null) fl.addAll(java.util.List.of(rowFonts[r]));
                rowFonts[r] = fl.toArray(new String[0]);
            }
        }
        colW = new double[maxCol];
        colScr = new double[maxCol];
        colW100 = new double[maxCol];
        rowH = new double[maxRow];
        rowEffPx = new double[maxRow];
        colX = new double[maxCol + 1];
        rowY = new double[maxRow + 1];
        setScale(1);
    }

    /** 印刷の換算モデル。 */
    public PrinterModel model() { return model; }

    /** 推定した作成者の画面 DPI (96 = 標準)。 */
    public int authorDpi() { return authorDpi; }

    /** 現在の倍率で、プリンタの横 1 ピクセルが何 pt (シート) か。 */
    public double devUnitX() { return 72.0 / hdpi / realScale(); }

    /** 指定フォントの文字の行送り (シートの pt): 現在の倍率での自動調整行の高さ。 */
    /** 縦書きの字送り (シートの pt): 倍率で実体化した自動調整行の高さ。 */
    public double stackPitch(String font, double sizePt) {
        if (font == null || sizePt <= 0) return 0;
        return model.autoRowDev(font, sizePt, realScale(), vdpi) * 72.0 / vdpi / realScale();
    }

    public double linePitch(String font, double sizePt) {
        if (font == null || sizePt <= 0) return 0;
        if (!model.authorEnv) return model.linePitchDev(font, sizePt, scale, vdpi) * 72.0 / vdpi / Math.max(0.01, scale);
        return model.autoRowDev(font, sizePt, realScale(), vdpi) * 72.0 / vdpi / realScale();
    }

    /** 行高 (pt) を画面ピクセルにする (端数のまま。9.15pt → 12.2 px)。 */
    static double toPx(double ht) {
        return ht / PX;
    }

    /**
     * 保存された既定の行高が、作成者の Excel の画面の標準行か (表示スケールにより自分の値と異なることがある)。
     * そうなら行の換算の基準にする。Excel の PDF で経験的に決めた: ＭＳ 系は x14ac:dyDescent が 0.15 以外で
     * 12.75〜13.4pt、游ゴシックは 17.5〜18.7pt。
     */
    /** 既定の行高が、作成者既定行の条件 ({@link #authorDefault}) に当たる値か (画面 DPI の指定によらない)。 */
    private boolean authorDefaultValue(double defPt, double dy) {
        if (model.ms) return defPt >= DEFROW_LO && defPt < 13.49 && (Math.abs(dy - 0.15) > 0.01);
        if (model.yu) return defPt >= 17.49 && defPt < 18.74;
        return false;
    }

    /**
     * 既定行を標準行にする ＭＳ 系の既定の行高の下限。
     */
    static final double DEFROW_LO = 12.5;

    private boolean authorDefault(double defPt, double dy) {
        // 印刷する PC の画面 DPI に 96 を指定したときは、Excel はその画面の標準行で換算する (保存された既定の行高は基準にしない。
        // 作成者既定行 13.2/13.0pt のブックで、手で決めた行が 0.33/0.50pt 高くなる)
        if (screenDpi == 96) return false;
        if (model.ms) {
            return defPt >= 12.74 && defPt < 13.49 && Math.abs(dy - 0.15) > 0.01;
        }
        if (model.yu) return defPt >= 17.49 && defPt < 18.74;
        return false;
    }

    /** 行の空でなく結合されていないセルの "フォント|サイズ" の一覧 (なければ null)。 */
    static final boolean ROWFMT_NOSTD = !"0".equals(System.getProperty("x2p.rowfmtnostd"));
    static final int CUSTOM_ROW_EXT = Integer.getInteger("x2p.customrowext", 0);
    static final int NOHT_REWRAP = Integer.getInteger("x2p.nohtrewrap", 1);
    static final boolean REWRAP_OUTER = !"0".equals(System.getProperty("x2p.rewrapouter"));

    /** レイアウトの列数 (maxCol) より右の列の幅 (文字数。非表示は -1)。col 要素になければ既定幅。 */
    private double outerColChars(int c) {
        for (CTCols cols : sheet.getCTWorksheet().getColsArray())
            for (CTCol cc : cols.getColArray())
                if (c + 1 >= cc.getMin() && c + 1 <= cc.getMax()) {
                    if (cc.isSetHidden() && cc.getHidden()) return -1;
                    if (cc.isSetWidth()) return cc.getWidth();
                }
        return defaultColChars;
    }

    /** 列の画面の幅 (px、96dpi)。非表示は 0。 */
    public double screenColPx(int c) {
        double chars = c < maxCol ? colChars[c] : outerColChars(c);
        if (chars <= 0 || (c < maxCol && colW[c] <= 0)) return 0;
        return Math.floor(((256 * chars + Math.floor(128.0 / mdw)) / 256) * mdw);
    }

    /**
     * 画面の寸法 (96dpi の列幅と文字の送り) で、左揃えのセルの文字が toCol 列の左端を越えてはみ出すか。
     * Excel は、印刷範囲のあるシートで次のページへ続くはみ出しを、画面の寸法で届くときだけ描く (夜間 n3507 は届く。
     * n2494・n3965 は印刷の列幅では届くが、画面の列幅では届かず、Excel は続きを描かない)。
     */
    public boolean screenReach(Cell cell, String text, int toCol) {
        if (text == null || text.isEmpty()) return false;
        XSSFFont f = ((XSSFCellStyle) cell.getCellStyle()).getFont();
        FontManager fm = model.fontManager();
        double size = f.getFontHeight() / 20.0;
        float pt = (float) (screenPpem(metricName(f), size) * 0.75);
        FontManager.Spec sp = fm.spec(metricName(f), f.getBold(), f.getItalic());
        String t = ComplexText.cluster(text);
        double w = 0;
        for (int i = 0; i < t.length(); ) {
            int cp = t.codePointAt(i);
            i += Character.charCount(cp);
            if (cp == '\n' || cp == '\r') break;
            w += Math.round(sp.charWidth(cp, pt, fm) / 0.75f);
        }
        double dist = 0;
        for (int c = cell.getColumnIndex(); c < toCol; c++) dist += screenColPx(c);
        if (System.getProperty("x2p.dbgspill") != null) System.err.println("REACH " + ((org.apache.poi.xssf.usermodel.XSSFCell) cell).getReference() + " -> " + toCol + " textPx=" + w + " distPx=" + dist + " printText=" + Math.round(sp.width(t, (float) size, fm)) + " printDist=" + Math.round(x(toCol) - x(cell.getColumnIndex())) + " font=" + metricName(f) + " " + size);
        return w + SCREEN_REACH_PAD > dist;
    }

    /** 画面の寸法のはみ出しの判定で、文字の幅に足す左の余白 (px)。 */
    static final double SCREEN_REACH_PAD = Double.parseDouble(System.getProperty("x2p.reachpad", "2"));

    /** REWRAP: 折り返して 2 行以上になるセルがあれば、行数 × 画面の標準行の最大 (px。なければ・表にないフォントがあれば 0)。 */
    private int rewrapPx(XSSFRow row) {
        FontManager fm = model.fontManager();
        int fit = 0;
        boolean multi = false;
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.STRING) continue;
            XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
            if (!cs.getWrapText() || mergedAt(cell.getRowIndex(), cell.getColumnIndex()) != null) continue;
            if (ROT_WRAP_KEEP && cs.getRotation() != 0) continue;   // 回転した文字の折り返しは横書きの行数で数えない (総当たり sw-rot: Excel の保存値のまま)
            if (cs.getAlignment() == org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER_SELECTION) continue;   // 列をまたぐ表示は数えない
            int c = cell.getColumnIndex();
            // 印刷する範囲の外の列のセルも行の高さに数える (夜間 n0628: 印刷範囲 A:X の外の AD 列の折り返しで 2 行分)
            double chars = c < maxCol ? colChars[c] : REWRAP_OUTER ? outerColChars(c) : 0;
            if (chars <= 0) continue;
            XSSFFont f = cs.getFont();
            double size = f.getFontHeight() / 20.0;
            int line = screenRow96(metricName(f), size);
            if (line <= 0) return 0;
            double px = Math.floor(((256 * chars + Math.floor(128.0 / mdw)) / 256) * mdw);
            XSSFRichTextString rs = (XSSFRichTextString) cell.getRichStringCellValue();
            if (rs.numFormattingRuns() > 1) {
                // 書式付き文字列: 折り返した行ごとに、その行の文字の画面の標準行の最大を足す
                List<TextLayout.Run> runs = new java.util.ArrayList<>();
                String all = rs.getString();
                int pos = 0;
                for (int i = 0; i < rs.numFormattingRuns(); i++) {
                    int st = rs.getIndexOfFormattingRun(i), en = Math.min(all.length(), st + rs.getLengthOfFormattingRun(i));
                    if (st > pos) runs.add(screenRun(all.substring(pos, st), f, fm));
                    XSSFFont rf = rs.getFontOfFormattingRun(i);
                    if (rf == null) rf = f;
                    if (en > st) runs.add(screenRun(all.substring(st, en), rf, fm));
                    pos = Math.max(pos, en);
                }
                if (pos < all.length()) runs.add(screenRun(all.substring(pos), f, fm));
                java.util.Map<String, Integer> px96 = new java.util.HashMap<>();
                for (TextLayout.Run ru : runs) {
                    int m = screenRow96(ru.fontName, ru.ptSize);
                    if (m <= 0) return 0;
                    px96.put(ru.fontName + "|" + ru.ptSize, m);
                }
                List<TextLayout.Line> lines = TextLayout.layoutCounting(runs, (float) ((px - REWRAP) * 0.75), fm);
                int h = 0;
                for (TextLayout.Line ln : lines) {
                    int lh = 0;
                    for (TextLayout.Run ru : ln.runs) {
                        Integer m = px96.get(ru.fontName + "|" + ru.ptSize);
                        if (m != null && !ru.text.isEmpty()) lh = Math.max(lh, m);
                    }
                    h += lh > 0 ? lh : line;
                }
                if (lines.size() >= 2) multi = true;
                fit = Math.max(fit, h);
                continue;
            }
            String t = cell.getStringCellValue().replace("\r", "");   // Excel は単独の CR (_x000D_) では改行しない
            TextLayout.Run run = new TextLayout.Run(t, fm.spec(metricName(f), f.getBold(), f.getItalic()),
                    (float) (screenPpem(metricName(f), size) * 0.75), java.awt.Color.BLACK);
            int n = TextLayout.layoutCounting(List.of(run), (float) ((px - REWRAP) * 0.75), fm).size();
            if (n >= 2) multi = true;
            fit = Math.max(fit, n * (line + Math.max(0, rubyPx(cell))));   // ふりがなは各行の上に 1 行分
        }
        if (!multi) return 0;
        int one = refitPx(row);
        if (one <= 0) return 0;
        return Math.min(546, Math.max(fit, one));   // Excel の行高の上限 409.5pt
    }

    /** REWRAP_RICH: 画面の大きさで測るラン (元のフォント名と大きさを fontName・ptSize に残す)。 */
    private TextLayout.Run screenRun(String text, XSSFFont f, FontManager fm) {
        double size = f.getFontHeight() / 20.0;
        TextLayout.Run r = new TextLayout.Run(text.replace("\r", ""), fm.spec(metricName(f), f.getBold(), f.getItalic()),
                (float) (screenPpem(metricName(f), size) * 0.75), java.awt.Color.BLACK);
        r.fontName = metricName(f);
        r.ptSize = size;
        return r;
    }

    /**
     * 「ふりがなを表示」(ph="1") の文字列のセルで、ふりがなのために各行の上に足す高さ (画面の px。ふりがなのフォント
     * (phoneticPr の fontId。なければセルのフォント) の標準行)。ふりがなを表示しないセルは 0。
     */
    /** 行にふりがなを表示するセル (結合セルを除く) があるか。 */
    private boolean hasRuby(XSSFRow row) {
        for (Cell cell : row)
            if (!merged.containsKey(key(cell.getRowIndex(), cell.getColumnIndex())) && rubyPx(cell) > 0) return true;
        return false;
    }

    static final boolean ROW_PH = "1".equals(System.getProperty("x2p.rowph"));   // 試験用 (列の phonetic で足すので既定は無効)

    /** シートのふりがなの書体 (phoneticPr の fontId。なければ styles の 2 番目のフォント) の字の高さ (画面の px)。 */
    private int sheetRubyPx() {
        XSSFFont pf = null;
        try {
            var ws = sheet.getCTWorksheet();
            int id = ws.isSetPhoneticPr() ? (int) ws.getPhoneticPr().getFontId() : 1;
            pf = sheet.getWorkbook().getStylesSource().getFontAt(id);
        } catch (RuntimeException ignored) {
            pf = null;
        }
        if (pf == null) return 0;
        FontManager.Spec sp = model.fontManager().spec(pf.getFontName(), false, false);
        long ppem = Math.round(pf.getFontHeight() / 20.0 * 4 / 3.0);
        return (int) Math.max(0, Math.floor((sp.ascent + sp.descent) * ppem + 1e-6));
    }

    private int rubyPx(Cell cell) {
        if (cell.getCellType() != CellType.STRING) return 0;
        var cc = ((org.apache.poi.xssf.usermodel.XSSFCell) cell).getCTCell();
        if (!cc.isSetPh() || !cc.getPh()) return 0;
        XSSFFont pf = null;
        var rs = (XSSFRichTextString) cell.getRichStringCellValue();
        var ct = rs == null ? null : rs.getCTRst();
        if (ct != null && ct.isSetPhoneticPr()) {
            try {
                pf = sheet.getWorkbook().getStylesSource().getFontAt((int) ct.getPhoneticPr().getFontId());
            } catch (RuntimeException ignored) {
                pf = null;
            }
        }
        if (pf == null) pf = ((XSSFCellStyle) cell.getCellStyle()).getFont();
        // ふりがなの行は字の高さだけ (標準行の行間を含まない): floor((ascent + descent) × 画面の字の大きさ)。
        // Excel の実測 (游ゴシック): 6pt で 12px、9pt で 19px
        FontManager.Spec sp = model.fontManager().spec(pf.getFontName(), false, false);
        long ppem = Math.round(pf.getFontHeight() / 20.0 * 4 / 3.0);
        return (int) Math.max(0, Math.floor((sp.ascent + sp.descent) * ppem + 1e-6));
    }

    /** 折り返しの数え直しで使う画面の字の大きさ (px)。 */
    /** 等幅の ＭＳ 明朝・ＭＳ ゴシックの + 1 は、画面の大きさが整数の px (12pt = 16px など) のときは付けない (夜間 n1110: 12pt の折り返しの行数)。 */
    static final int PPEM_INT = 1;

    private static double screenPpem(String font, double size) {
        long p = Math.round(size * 4 / 3.0);
        if (!fixedMs(font)) return p;
        if (PPEM_INT == 1 && Math.abs(size * 4 / 3.0 - p) < 1e-6) return p;   // 画面の大きさが整数の px なら + 1 しない
        if (PPEM_INT == 2) return p;
        return p + 1;
    }

    /** ＭＳ 明朝・ＭＳ ゴシックと HGｺﾞｼｯｸ・HG明朝 (固定幅。Ｐ・Ｓ 付きと HG丸ｺﾞｼｯｸ を除く。Fableレポート D: 試験 fb-t2 の wrap 156 件)。 */
    private static boolean fixedMs(String name) {
        String n = FontManager.norm(name);
        if (n.matches("hg(ｺﾞｼｯｸ|明朝).*")) return true;
        return n.startsWith("ms ") && !n.startsWith("ms p");
    }

    /**
     * 変換する環境の画面 (96 dpi) の標準行 (px)。表にないフォントは 0。
     */
    private static int screenRow96(String font, double size) {
        int m = PrinterModel.measuredScreenRow(font, size, 96);
        return m;
    }

    /** HTONLY_REFIT: 行の全セル (複数行の結合を除く) と行の書式のフォントの、変換する環境の画面の標準行の最大 (px。表にないフォントがあれば 0)。 */
    /** シートの標準行を決めるフォント「名前|大きさ」: 列の書式のフォント (16384 列目までのうち書式のない列があれば標準フォントも)。 */
    public Set<String> sheetRowFonts() {
        Set<String> fs = new LinkedHashSet<>();
        long covered = 0;
        for (CTCols cols : sheet.getCTWorksheet().getColsArray())
            for (CTCol c : cols.getColArray()) {
                if (!c.isSetStyle()) continue;
                covered += c.getMax() - c.getMin() + 1;
                XSSFFont f = sheet.getWorkbook().getCellStyleAt((int) c.getStyle()).getFont();
                fs.add(fkey(f));
            }
        if (covered < 16384) fs.add(model.defaultFont() + "|" + model.defaultSize());
        return fs;
    }

    /** COLSTD: {@link #sheetRowFonts} の画面の標準行の最大 (px。表にないフォントがあれば 0)。 */
    private int colStdPx(CTWorksheet ws) {
        Set<String> fs = sheetRowFonts();
        int fit = 0;
        for (String k : fs) {
            int bar = k.lastIndexOf('|');
            int m = screenRow96(k.substring(0, bar), Double.parseDouble(k.substring(bar + 1)));
            if (m <= 0) return 0;
            fit = Math.max(fit, m);
        }
        // ふりがなを表示する列 (col の phonetic="1") があれば、ふりがなの 1 行分を足す (夜間 n3716: 30 + 13 = 43px = 32.25pt。
        // row 要素のない行も、セルのない高さのない行もこの高さ)
        if (fit > 0 && COL_PH) {
            for (CTCols cols : ws.getColsArray())
                for (CTCol c : cols.getColArray())
                    if (c.isSetPhonetic() && c.getPhonetic() && !(c.isSetHidden() && c.getHidden())) { colPhRuby = true; return fit + sheetRubyPx(); }
        }
        return fit;
    }
    static final boolean COL_PH = !"0".equals(System.getProperty("x2p.colph"));
    /** 列の標準行にふりがなの分を足したか (足したら、高さのない行でセルのふりがなを重ねて足さない)。 */
    private boolean colPhRuby;

    /**
     * Excel の行の自動調整: 行の全列について、セルがあればセルのフォント、なければ行の書式、なければ列の書式、なければ標準フォントの、
     * 画面の標準行の最大 (px。表にないフォントがあれば 0)。複数行の結合セルは数えない。太い罫線の印ごとに +1。
     */
    /** autoRowPx で値のある 1 行の結合セルを数えないか (ht のある行の自動調整の下限のときだけ)。 */
    private boolean skipMerged;
    static final boolean AUTO_SKIP_MERGE = true;
    /** 回転した文字の折り返しのセルを、行の数え直し (rewrapPx・hasWrapText) に数えない。 */
    static final boolean SEL_WRAP_KEEP = !"0".equals(System.getProperty("x2p.selwrap"));
    static final boolean ROT_WRAP_KEEP = !"0".equals(System.getProperty("x2p.rotwrap"));
    static final boolean BASELINE_SUM = true;
    static final boolean PLAIN_KEEP = true;
    /** autoRowPx の書体の標準行の最大 (基準線の和を足す前。px)。 */
    private int lastPlainFit;

    private static boolean cellEmpty(Cell c) {
        return c.getCellType() == CellType.BLANK || c.getCellType() == CellType.STRING && c.getStringCellValue().isEmpty();
    }

    private int autoRowPx(XSSFRow row) {
        Map<String, Boolean> fs = new java.util.LinkedHashMap<>();
        Set<String> texts = new java.util.HashSet<>(); // 文字のある下揃えのセルのフォント
        boolean uncovered = false; // 書式のない列があるか
        Set<Integer> have = new java.util.HashSet<>();
        for (Cell cell : row) {
            have.add(cell.getColumnIndex());
            CellRangeAddress mg = merged.get(key(cell.getRowIndex(), cell.getColumnIndex()));
            if (mg != null && mg.getFirstRow() != mg.getLastRow()) continue;
            if (skipMerged && mg != null && !cellEmpty(cell)) continue;   // 値のある 1 行の結合セルの字で高くしない (夜間 n0902)
            XSSFCellStyle st = (XSSFCellStyle) cell.getCellStyle();
            if (st.getRotation() != 0) continue;
            addAutoFont(fs, st);
            if (cell.getCellType() != org.apache.poi.ss.usermodel.CellType.BLANK && st.getVerticalAlignment() == org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM)
                texts.add(fkey(st.getFont()));
            // 書式付き文字列は、文字のフォントも数える
            if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING) {
                org.apache.poi.xssf.usermodel.XSSFRichTextString rs = (org.apache.poi.xssf.usermodel.XSSFRichTextString) cell.getRichStringCellValue();
                boolean bottom = st.getVerticalAlignment() == org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM;
                for (int i = 0; i < rs.numFormattingRuns(); i++) {
                    XSSFFont rf = rs.getFontOfFormattingRun(i);
                    if (rf == null) continue;
                    String k = fkey(rf);
                    fs.merge(k, bottom, Boolean::logicalOr);
                    if (bottom) texts.add(k);
                }
            }
        }
        if (row.isFormatted() && row.getRowStyle() != null) {
            addAutoFont(fs, (XSSFCellStyle) row.getRowStyle());
        } else {
            long covered = 0;
            for (CTCols cols : sheet.getCTWorksheet().getColsArray())
                for (CTCol c : cols.getColArray()) {
                    if (!c.isSetStyle()) continue;
                    long n = c.getMax() - c.getMin() + 1;
                    covered += n;
                    long inCells = 0;
                    for (int col : have) if (col + 1 >= c.getMin() && col + 1 <= c.getMax()) inCells++;
                    if (inCells < n) addAutoFont(fs, sheet.getWorkbook().getCellStyleAt((int) c.getStyle()));
                }
            if (covered < 16384) {
                fs.merge(model.defaultFont() + "|" + model.defaultSize(), true, Boolean::logicalOr);
                uncovered = true;
            }
        }
        int fit = 0, above = 0, below = 0;
        boolean parts = true;
        for (var e : fs.entrySet()) {
            String k = e.getKey();
            int bar = k.lastIndexOf('|');
            double size = Double.parseDouble(k.substring(bar + 1));
            int m = screenRow96(k.substring(0, bar), size);
            if (m <= 0) return 0;
            fit = Math.max(fit, m);
            boolean text = texts.contains(k), std = uncovered && k.equals(model.defaultFont() + "|" + model.defaultSize());
            if (!text && !std) continue;
            int[] p = FontCalc.rowParts(k.substring(0, bar), size, 96);
            // 計算できないフォント、計算が実測表と違うフォント (Tahoma など) があれば、標準行の最大にする
            if (p == null || p[0] + p[1] != m) { parts = false; continue; }
            // 文字のある下揃えのセルと、書式のない列 (標準フォント) は基準線がそろう: 上の最大 + 下の最大
            if (text) above = Math.max(above, m - p[1]);
            below = Math.max(below, p[1]);
        }
        lastPlainFit = fit;
        if (parts && BASELINE_SUM) fit = Math.max(fit, above + below);
        // ふりがなを表示するセル (結合セルを除く): セルのフォントの標準行 + ふりがなのフォントの標準行
        for (Cell cell : row) {
            if (merged.containsKey(key(cell.getRowIndex(), cell.getColumnIndex()))) continue;   // 結合セルは数えない
            int ru = rubyPx(cell);
            if (ru <= 0) continue;
            XSSFFont f = ((XSSFCellStyle) cell.getCellStyle()).getFont();
            int m = screenRow96(metricName(f), f.getFontHeight() / 20.0);
            if (m > 0) fit = Math.max(fit, m + ru);
        }
        if (row.getCTRow().getThickTop()) fit++;
        if (row.getCTRow().getThickBot()) fit++;
        return fit;
    }

    /** 行の高さを決める書体: テーマのフォント (scheme) なら、日本語版の Excel ではテーマの日本語の書体 (exp55: Calibri → ＭＳ Ｐゴシック)。 */
    private String metricName(XSSFFont f) {
        org.apache.poi.ss.usermodel.FontScheme s = f.getScheme();
        if (s == org.apache.poi.ss.usermodel.FontScheme.MINOR && model.minorJpan != null) return model.minorJpan;
        if (s == org.apache.poi.ss.usermodel.FontScheme.MAJOR && model.majorJpan != null) return model.majorJpan;
        // Excel 以外で保存したブックの、寸法の分からない書体 (Linux の LibreOffice が書く Noto Sans CJK SC など。Windows にない) は、
        // 標準の書体の寸法で測る (exp111)
        if (model.nonExcel && PrinterModel.measuredScreenRow(f.getFontName(), f.getFontHeight() / 20.0, 96) <= 0) return model.defaultFont();
        return f.getFontName();
    }

    private String fkey(XSSFFont f) { return metricName(f) + "|" + f.getFontHeight() / 20.0; }

    private void addAutoFont(Map<String, Boolean> fs, XSSFCellStyle st) {
        XSSFFont f = st.getFont();
        fs.merge(fkey(f), st.getVerticalAlignment() == org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM, Boolean::logicalOr);
    }

    private int refitPx(XSSFRow row) { return refitPx(row, false, false); }

    /**
     * @param withDefault 標準フォントも必ず数える
     * @param noMerged    結合セルを数えない
     */
    private int refitPx(XSSFRow row, boolean withDefault, boolean noMerged) {
        Set<String> fs = new LinkedHashSet<>();
        if (withDefault || row.getPhysicalNumberOfCells() == 0) fs.add(model.defaultFont() + "|" + model.defaultSize());
        if (row.isFormatted() && row.getRowStyle() != null) {
            XSSFFont rf = ((XSSFCellStyle) row.getRowStyle()).getFont();
            fs.add(fkey(rf));
        }
        for (Cell cell : row) {
            CellRangeAddress mg = merged.get(key(cell.getRowIndex(), cell.getColumnIndex()));
            if (mg != null && (noMerged || mg.getFirstRow() != mg.getLastRow())) continue;
            if (((XSSFCellStyle) cell.getCellStyle()).getRotation() != 0) continue;
            XSSFFont f = ((XSSFCellStyle) cell.getCellStyle()).getFont();
            fs.add(fkey(f));
        }
        int fit = 0;
        for (String k : fs) {
            int bar = k.lastIndexOf('|');
            int m = screenRow96(k.substring(0, bar), Double.parseDouble(k.substring(bar + 1)));
            if (m <= 0) return 0;
            fit = Math.max(fit, m);
        }
        // 太い罫線の印 (thickTop / thickBot) がある行は 1 px 高い
        if (row.getCTRow().getThickTop()) fit++;
        if (row.getCTRow().getThickBot()) fit++;
        return fit;
    }

    private String[] fontsOf(XSSFRow row) {
        Set<String> out = new LinkedHashSet<>();
        for (Cell cell : row) {
            if (CellPainter.isEmpty((XSSFCell) cell)) continue;
            CellRangeAddress mg = merged.get(key(cell.getRowIndex(), cell.getColumnIndex()));
            if (mg != null) continue;
            XSSFFont f = ((XSSFCellStyle) cell.getCellStyle()).getFont();
            out.add(fkey(f));
        }
        return out.isEmpty() ? null : out.toArray(new String[0]);
    }

    /**
     * 作成者の画面 DPI を推定する。列幅は作成者の最大数字幅の 1/256 で切り捨てた文字数で保存されるため、その端数から
     * 数字幅が分かる: 8 px なら 96 dpi、9 px なら 120 dpi、11 px なら 144 dpi (ＭＳ Ｐゴシック 11)。Excel はシートを
     * 実行している PC の画面の値で印刷するので、これらのファイルでは高 DPI の作成者の PC の値になる: 行・列はその画面の
     * ピクセルから換算する (実測: 行は 120 / 144 dpi で 2.3% / 3.9% 高い)。
     */
    private void detectAuthorDpi(CTWorksheet ws, CTSheetFormatPr fp) {
        // 游ゴシックも対象 (dyDescent は 96 dpi で 0.4、120 dpi で 0.45、144 dpi で 0.55 だが、その印のないファイル
        // にも効くので条件にはしない)
        boolean forced = screenDpi > 0;
        // 画面 DPI に 96 を指定したときは推定しない (自動調整のやり直しも働かない)
        if (forced && screenDpi == 96) return;
        if (!forced) {
            if (!(model.ms || model.yu)) return;
        }
        Set<Double> widths = new TreeSet<>();
        for (CTCols cols : ws.getColsArray())
            for (CTCol c : cols.getColArray()) if (c.isSetWidth()) widths.add(c.getWidth());
        if (fp != null && fp.isSetDefaultColWidth()) widths.add(fp.getDefaultColWidth());
        widths.removeIf(w -> w <= 0 || Math.abs(w - Math.rint(w)) < 1e-9);
        int best = 96, bestN = consistent(widths, mdw);
        if (widths.isEmpty() && !forced) {
            // 列幅の指定がない (すべて既定の幅): 既定の行高が高 DPI の画面の標準行と一致すれば、その DPI とする
            best = defaultRowDpi(fp);
            if (best == 96) return;
            bestN = 0;
        }
        for (int d : new int[] {120, 144, 168, 192}) {
            if (widths.isEmpty()) break;
            int m = (int) Math.max(1, Math.round(model.digitEm * Math.round(model.defaultSize() * d / 72.0)));
            int mt = PrinterModel.measuredMdw(model.defaultFont(), model.defaultSize(), d);
            if (mt > 0) m = mt;
            int k = consistent(widths, m);
            if (k > bestN) { best = d; bestN = k; }
        }
        // 列幅から推定した作成者 DPI (画面 DPI を指定したときに、印刷する PC と作成者の DPI が違うかの判定に使う)
        inferredDpi = !widths.isEmpty() && bestN >= widths.size() ? best : 96;
        if (forced) { best = screenDpi; bestN = widths.size(); }
        if (best == 96 || bestN < widths.size()) return;
        authorDpi = best;
        long pp = Math.round(model.defaultSize() * best / 72.0);
        authorMdw = (int) Math.max(1, Math.round(model.digitEm * pp));
        int mm = PrinterModel.measuredMdw(model.defaultFont(), model.defaultSize(), best);
        if (mm > 0) authorMdw = mm;
        // 游ゴシック 11pt: 高 DPI の画面の標準行は 18pt (120 dpi で 30 px、144 dpi で 36 px)
        authorRow = model.yu ? (int) Math.round(18 * model.defaultSize() / 11.0 * best / 72.0)
                : PrinterModel.measuredScreenRow(model.defaultFont(), model.defaultSize(), best) > 0
                ? PrinterModel.measuredScreenRow(model.defaultFont(), model.defaultSize(), best)
                : model.screenRowAt(pp);
    }

    /**
     * 既定の行高 (ユーザー指定でないもの) が、96 dpi 以外の画面の標準行 (ＭＳ 系) とちょうど一致すればその DPI、
     * そうでなければ 96 を返す。行の補正 ({@link #authorDefault}) と同じ条件のファイルに限る。
     */
    private int defaultRowDpi(CTSheetFormatPr fp) {
        if (!model.ms || fp == null || fp.getDefaultRowHeight() <= 0) return 96;
        if (fp.isSetCustomHeight() && fp.getCustomHeight()) return 96;
        double def = fp.getDefaultRowHeight();
        if (!authorDefault(def, dyDescent)) return 96;
        for (int d : new int[] {120, 144, 168, 192}) {
            long p = Math.round(model.defaultSize() * d / 72.0);
            if (Math.abs(model.screenRowAt(p) * 72.0 / d - def) < 0.01) return d;
        }
        return 96;
    }

    /** 数字幅 m のとき、1/256 の切り捨てで説明できる列幅 (文字数) の数。 */
    private static int consistent(Set<Double> widths, int m) {
        int k = 0;
        for (double w : widths) {
            double v = w * m;
            if (Math.abs(v - Math.rint(v)) <= m / 256.0 + 1e-6) k++;
        }
        return k;
    }

    /** 印刷倍率 (1 = 100%) に合わせて列幅・行高を計算し直す。 */
    public void setScale(double scale) {
        this.scale = scale;
        scale = realScale();
        PrinterModel.Factors f = model.factors(scale, hdpi, vdpi);
        if (model.yu) {
            // 游ゴシック: アウトラインの数字幅 (1139/2048 × ppem) を丸めずに使う
            long px = Math.round(model.defaultSize() * scale * hdpi / 72.0);
            double od = FontManager.outlineDigit(model.defaultFont());
            if (od > 0) f.mdwP = Math.max(1, Math.round(model.digitEm * px));
        }
        int cm = mdw;  // 列幅の画面ピクセルを求める数字幅
        if (authorDpi != 96) {
            // 保存された高さは作成者の画面ピクセル: プリンタの標準行 / 作成者の標準行 で換算する。列も作成者の数字幅
            f.rowRatio = f.rowRatio * model.screenRow / authorRow * authorDpi / 96.0;
            cm = authorMdw;
        } else if (!defaultCustomHeight && authorDefault(defaultRowPx * PX, dyDescent)) {
            // シートの既定の行高が、作成者の Excel の画面の標準行
            f.rowRatio = f.rowRatio * model.screenRow / defaultRowPx;
        }
        for (int i = 0; i < maxCol; i++) {
            double w = colChars[i];
            if (w < 0) { colW[i] = 0; colW100[i] = 0; continue; }
            // Excel の式で画面ピクセルを求め、プリンタのピクセルに換算する (偶数丸め)
            double px = Math.floor(((256 * w + Math.floor(128.0 / cm)) / 256) * cm);
            if (showFormulas) px *= 2;   // 数式の表示のシートは、Excel が列幅を 2 倍にして印刷する (夜間 n1375 など)
            double dev = Math.floor(px * f.mdwP / cm + 0.5);
            double dev100 = dev;
            if (linear && this.scale != 1)
                dev = Math.floor(dev * this.scale + 0.5 + 1e-9) / this.scale;   // 100% のプリンタの画素 × 倍率を列ごとに四捨五入 (拡大も同じ。exp103 の 130%)
            if (numWidePx != null && i < numWidePx.length && numWidePx[i] >= 0) {
                if (numWidePx[i] > px) dev = Math.rint(numWidePx[i] * f.mdwP / cm);
                else dev += NUM_WIDEN_ADD;
                if (linear && this.scale != 1) dev = Math.floor(dev * this.scale + 0.5 + 1e-9) / this.scale;
            }
            colW[i] = dev * f.ptX;
            colW100[i] = dev100 * f.ptX;   // Fableレポート B の試験と同じく、長い数値の広がりは含めない
            if (fitPx != null && fitPx[i] > px && fitPx[i] > (numWidePx != null && i < numWidePx.length ? numWidePx[i] : 0)) {
                dev = Math.rint(fitPx[i] * f.mdwP / cm);
                if (linear && this.scale != 1) dev = Math.floor(dev * this.scale + 0.5 + 1e-9) / this.scale;
                colW[i] = dev * f.ptX;
            }
            colScr[i] = (numWidePx != null && i < numWidePx.length && numWidePx[i] > px ? numWidePx[i] : px) * PX;
        }
        for (int i = 0; i < maxRow; i++) {
            // 行ごとにデバイスピクセルへ切り捨てる (実測で確認済み)
            double rp = rowPx[i];
            double dev = Math.floor(rp * f.rowRatio + 1e-9);
            if (!model.authorEnv) dev = Math.floor(Math.rint(rp * 4) / 4 * f.rowRatio + 0.125 + 1e-9);
            if (linear && this.scale != 1) dev = Math.floor(dev * this.scale + 0.5 + 1e-9) / this.scale;
            // 自動調整し直すのは、印刷する PC の画面 DPI を指定し、それが作成者の DPI と違うときだけ (同じ DPI なら保存値が Excel 自身の
            // 自動調整の結果で、こちらの計算より正確。常には使わない)
            if (rowHtOnly[i] && rowPx[i] > 0 && screenDpi > 0 && (screenDpi != inferredDpi)) {
                // Excel は変換する環境の画面で自動調整した高さ (各フォントの画面の標準行の最大) を、固定した行と同じ比で印刷に換算する
                int d = authorDpi;
                double stdAuthorPx96 = d != 96 ? authorRow * 96.0 / d : model.screenRow;
                if (rowPx[i] <= stdAuthorPx96 + 1.01) {
                    String[] fs = rowFonts[i];
                    int fit = 0;
                    if (fs == null) fit = d != 96 ? authorRow : model.screenRow;
                    else for (String k : fs) {
                        int bar = k.lastIndexOf('|');
                        int m = PrinterModel.measuredScreenRow(k.substring(0, bar), Double.parseDouble(k.substring(bar + 1)), d);
                        if (m <= 0) { fit = 0; break; }   // 表にないフォントがあれば保存値のまま
                        fit = Math.max(fit, m);
                    }
                    if (fit > 0) dev = Math.floor(fit * 96.0 / d * f.rowRatio + 1e-9);
                }
            }
            rowH[i] = dev * f.ptY;
            rowEffPx[i] = rowPx[i] > 0 && dev != Math.floor(rowPx[i] * f.rowRatio + 1e-9) ? (dev + 0.5) / f.rowRatio : rowPx[i];
        }
        defaultRowH = (!model.authorEnv ? Math.floor(Math.rint(defaultRowPx * 4) / 4 * f.rowRatio + 0.125 + 1e-9)
                : Math.floor(defaultRowPx * f.rowRatio + 1e-9)) * f.ptY;  // 行と同じく誤差で 1 px 減らないように
        for (int i = 0; i < maxCol; i++) colX[i + 1] = colX[i] + colW[i];
        for (int i = 0; i < maxRow; i++) rowY[i + 1] = rowY[i] + rowH[i];
        xScale = f.mdwP / mdw * f.ptX / PX;
        yScale = f.rowRatio * f.ptY / PX;
    }

    /** セルのキー (行 × 16384 + 列)。 */
    static long key(int r, int c) { return (long) r * 16384 + c; }

    /** セルを含む結合範囲 (なければ null)。 */
    public CellRangeAddress mergedAt(int r, int c) { return merged.get(key(r, c)); }

    /** 列の左端の位置 (シートの pt)。 */
    public double x(int col) {
        if (col <= 0) return 0;
        if (col >= maxCol) return colX[maxCol] + (col - maxCol) * colW[maxCol - 1];
        return colX[col];
    }

    /** 行の上端の位置 (シートの pt)。 */
    public double y(int row) {
        if (row <= 0) return 0;
        if (row >= maxRow) return rowY[maxRow] + (row - maxRow) * defaultRowH;
        return rowY[row];
    }

    /** 列幅 (シートの pt)。 */
    public double cw(int c) { return c < maxCol ? colW[c] : colW[maxCol - 1]; }

    /** 倍率で丸める前 (100%) の列幅 (シートの pt)。 */
    public double cw100(int c) { return c < maxCol ? colW100[c] : colW100[maxCol - 1]; }

    /** 画面の列幅 (pt、プリンタへの換算前。Excel の Columns.Width)。 */
    public double screenColPt(int c) { return c < maxCol ? colScr[c] : colScr[maxCol - 1]; }

    /** 画面の行高 (pt、プリンタへの換算前)。 */
    public double screenRowPt(int r) { return r < maxRow ? rowEffPx[r] * PX : defaultRowPx * PX; }

    /** 自動調整し直した後の画面の行高 (px、プリンタのピクセルから逆算)。 */
    private double[] rowEffPx;

    /** 行高 (シートの pt)。 */
    public double rh(int r) { return r < maxRow ? rowH[r] : defaultRowH; }

    /**
     * 値または見える書式のあるセルの範囲を返す。
     *
     * @return {先頭行, 先頭列, 最終行, 最終列}。何もなければ null。先頭は常に A1
     */
    /** 使用範囲で非表示の列のセルを数えないか。 */
    static final boolean HIDDEN_SKIP = true;

    public static int[] usedRange(XSSFSheet sheet, boolean allMerged) {
        // 印刷範囲がなければ Excel は A1 から印刷する。終わりは値か罫線・塗りつぶしのあるセル
        // 非表示の列のセルは数えない (何も印刷されない。夜間 n0467: 非表示の N 列の文字で K〜M の空き列まで印刷していた)
        int lr = -1, lc = -1;
        for (var row : sheet) {
            for (Cell cell : row) {
                if (!significant(cell)) continue;
                if (HIDDEN_SKIP && sheet.isColumnHidden(cell.getColumnIndex())) continue;
                lr = Math.max(lr, cell.getRowIndex()); lc = Math.max(lc, cell.getColumnIndex());
            }
        }
        for (CellRangeAddress a : sheet.getMergedRegions()) {
            Cell cell = sheet.getRow(a.getFirstRow()) == null ? null : sheet.getRow(a.getFirstRow()).getCell(a.getFirstColumn());
            if (!allMerged && (cell == null || !significant(cell))) continue;
            lr = Math.max(lr, a.getLastRow()); lc = Math.max(lc, a.getLastColumn());
        }
        // スタイルのあるテーブルは、値のないセルまで印刷する
        for (var t : sheet.getTables()) {
            if (!TableStyles.styled(t) || t.getCellReferences() == null) continue;
            lr = Math.max(lr, t.getCellReferences().getLastCell().getRow());
            lc = Math.max(lc, t.getCellReferences().getLastCell().getCol());
        }
        if (lr < 0) return null;
        if (CUSTOM_ROW_EXT > 0) {
            // 高さを手で決めた行 (customHeight) は、セルがなくても Excel の使用範囲に入る (夜間 n3288: 末尾の 2 行で縦の中央寄せが 20pt 変わる)
            int ext = lr;
            var fpr = sheet.getCTWorksheet().getSheetFormatPr();
            boolean defCustom = fpr != null && fpr.isSetCustomHeight() && fpr.getCustomHeight();
            for (var row : sheet) {
                var ct = ((org.apache.poi.xssf.usermodel.XSSFRow) row).getCTRow();
                if (row.getRowNum() > ext && ct.isSetHt() && ct.isSetCustomHeight() && ct.getCustomHeight() && !row.getZeroHeight()
                        && (CUSTOM_ROW_EXT == 1 || CUSTOM_ROW_EXT == 2 && row.getPhysicalNumberOfCells() == 0
                            || CUSTOM_ROW_EXT == 3 && defCustom || CUSTOM_ROW_EXT == 4 && defCustom && row.getPhysicalNumberOfCells() == 0)) ext = row.getRowNum();
            }
            lr = ext;
        }
        if (COL_STYLE_ROWS > 0) {
            // セルのない行 (row 要素だけ) は、列の書式を継ぐ。列の書式に塗りつぶし・罫線があれば、Excel はその行までを印刷する
            // (夜間 n3288: 全列が白の塗りの書式で、末尾の 2 行は row 要素だけ。縦の中央寄せが 20pt 変わる)
            boolean[] colSig = new boolean[lc + 1];
            boolean any = false;
            for (int c = 0; c <= lc; c++) {
                if (sheet.isColumnHidden(c)) continue;
                colSig[c] = styleSignificant(sheet.getColumnStyle(c));
                any |= colSig[c];
            }
            int ext = lr;
            for (var row : sheet) {
                if (row.getRowNum() <= ext || row.getZeroHeight()) continue;
                var xr = (org.apache.poi.xssf.usermodel.XSSFRow) row;
                if (COL_STYLE_ROWS == 2 && !(xr.getCTRow().isSetCustomHeight() && xr.getCTRow().getCustomHeight())) continue;
                boolean rowFmt = row.isFormatted();
                boolean hit = false;
                if (rowFmt) hit = styleSignificant(row.getRowStyle()) && row.getPhysicalNumberOfCells() == 0;
                else if (any) for (int c = 0; c <= lc && !hit; c++) hit = colSig[c] && row.getCell(c) == null;
                if (hit) ext = row.getRowNum();
            }
            lr = ext;
        }
        return new int[] {0, 0, lr, lc};
    }

    /** 広がるシートの、幅の指定のない A〜AA 列に足す画素。 */
    private static final int NUM_WIDEN_ADD = 6;
    /** col 要素のないシートは長い数値でも列を広げない。 */
    static final boolean NW_NEED_COLS = true;
    /** 表示形式のある数値も列を広げる (exp110)。 */
    static final boolean NUM_WIDEN_FMT = true;
    /** 表示形式のある数値で広げるきっかけの文字数。 */
    static final int NUM_WIDEN_FMT_MIN = 8;
    /** 表示形式のある数値は 表示の文字数 − 1 で表を引く (exp110 の 76・84・93 px に近い)。 */
    static final int NWF_OFF = 1;
    private static final org.apache.poi.ss.usermodel.DataFormatter NUM_DF = new org.apache.poi.ss.usermodel.DataFormatter(java.util.Locale.ROOT);
    /** 広がる対象の最後の列 (AA)。 */
    private static final int NUM_WIDEN_LAST = 26;
    /** 数字幅・ppem ごとの、6〜11 桁の数値の列の幅 (画面 px)。 */
    private static final java.util.Map<Integer, int[]> NUM_WIDEN_TABLE = loadNumWiden();
    /** 広がるシートの列の幅 (画面 px。0 は既定幅のまま 6 画素足す列、-1 は対象外)。広がらないシートは null。 */
    private int[] numWidePx;

    /** 長い数値で列を広げた列か (numWiden)。 */
    public boolean numWidened(int c) { return numWidePx != null && c >= 0 && c < numWidePx.length && numWidePx[c] > 0; }

    private static java.util.Map<Integer, int[]> loadNumWiden() {
        java.util.Map<Integer, int[]> m = new java.util.TreeMap<>();
        try (var in = SheetLayout.class.getResourceAsStream("/numwiden.tsv")) {
            if (in == null) return m;
            for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\n")) {
                if (l.isBlank() || l.startsWith("#")) continue;
                String[] p = l.trim().split("\t");
                int[] w = new int[6];
                for (int i = 0; i < 6; i++) w[i] = Integer.parseInt(p[i + 2]);
                m.put(Integer.parseInt(p[0]) * 1000 + Integer.parseInt(p[1]), w);
            }
        } catch (java.io.IOException e) {
            // 表なし: 広げない
        }
        return m;
    }

    /**
     * セルのフォントの n 桁 (6〜11) の数値の列の幅 (画面 px)。式は未解明で、実測の表による近似。表のキーがなければ、
     * 20 ppem 以上は「ppem × 2/3 × 桁 + 10」、それ未満は同じ数字幅で一番近い ppem の値を使う。標準フォントの数字幅への依存は入れていない。
     */
    private int numWidth(XSSFFont font, int n) {
        double size = font.getFontHeight() / 20.0;
        int ppem = (int) Math.floor(size * 96 / 72 + 0.5);
        int d = PrinterModel.measuredMdw(font.getFontName(), size, 96);
        if (d <= 0) d = (int) Math.round(ppem * model.digitEm);
        int[] exact = NUM_WIDEN_TABLE.get(d * 1000 + ppem);
        if (exact != null) return exact[Math.min(n, 11) - 6];
        if (ppem >= 20) return (int) Math.round(ppem * 2.0 / 3 * n + 10);
        int best = -1;
        long score = Long.MAX_VALUE;
        for (int k : NUM_WIDEN_TABLE.keySet()) {
            long v = Math.abs(k / 1000 - d) * 1000L + Math.abs(k % 1000 - ppem);
            if (v < score) { score = v; best = k; }
        }
        if (best < 0) return 0;
        return (int) Math.round(NUM_WIDEN_TABLE.get(best)[Math.min(n, 11) - 6] * (double) d / (best / 1000));
    }

    /**
     * <cols> のあるシートで、幅の指定のない A〜AA 列に標準の表示形式の長い数値 (8 文字以上) があると、Excel は印刷で列を広げる。
     * このとき幅の指定のない A〜AA 列はすべて 6 画素広がり、7 桁以上の数値の列はさらに広がる。Excel の動きは実行により変わるが、多くの出力で広がる。
     */
    private void numWiden(boolean[] widthSet) {
        int[] wide = new int[Math.min(maxCol, NUM_WIDEN_LAST + 1)];
        boolean found = false;
        // 列の幅の指定 (col 要素) が実際にあるか (POI は読むときに空の cols を作るので、要素の数で見る)
        boolean hasColElems = false;
        for (CTCols cs : sheet.getCTWorksheet().getColsArray()) if (cs.sizeOfColArray() > 0) hasColElems = true;
        if (NW_NEED_COLS && !hasColElems) return;   // 列の幅の指定 (col 要素) がないシートは広げない (夜間 n0425)
        for (var row : sheet) {
            for (Cell cell : row) {
                int c = cell.getColumnIndex();
                if (c >= wide.length || widthSet[c] || colChars[c] < 0) continue;
                CellType t = cell.getCellType() == CellType.FORMULA ? cell.getCachedFormulaResultType() : cell.getCellType();
                if (t != CellType.NUMERIC) continue;
                if (cell.getCellStyle().getDataFormat() != 0) {
                    // 表示形式のある数値: 表示の文字数で同じ表を引く (Apache POI で作ったブック exp110: 「#,##0.00」の 8〜10 文字 → 76・84・93 px)
                    if (!NUM_WIDEN_FMT || !hasColElems || !model.nonExcel) continue;   // Excel で保存したブックは広げない (exp9・16・20)
                    String txt = NUM_DF.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString());
                    int len = txt.length();
                    if (len >= NUM_WIDEN_FMT_MIN) found = true;
                    if (len < 7 || txt.indexOf('#') >= 0) continue;
                    wide[c] = Math.max(wide[c], numWidth((XSSFFont) ((XSSFCellStyle) cell.getCellStyle()).getFont(), Math.max(6, Math.min(len - NWF_OFF, 11))));
                    continue;
                }
                double v = cell.getNumericCellValue(), a = Math.abs(v);
                if (a < 1) continue;
                int digits = (int) Math.floor(Math.log10(a)) + 1;
                boolean neg = v < 0, frac = a != Math.rint(a) && digits < 11;
                if (digits + (neg ? 1 : 0) >= 8) found = true;
                if (digits < 7) continue;
                int px = numWidth((XSSFFont) ((XSSFCellStyle) cell.getCellStyle()).getFont(), Math.min(digits + (frac ? 2 : 0), 11)) + (neg ? 5 : 0);
                wide[c] = Math.max(wide[c], px);
            }
        }
        if (!found) return;
        for (int c = 0; c < wide.length; c++) if (widthSet[c] || colChars[c] < 0) wide[c] = -1;
        numWidePx = wide;
    }

    /**
     * VALUE_FIT: 印刷のときに、値の入る幅まで広げる列の画面の幅 (px)。0 は広げない。
     * Excel は、customWidth の列が 1 つでもあるシートで、customWidth のない列 (全列 1〜16384 の customWidth の指定の中の列も含む) に
     * 入らない真偽値・エラー値・整数があると、印刷のときだけその列を「文字の幅 + 約 2 文字」に広げる (画面の列幅は変わらない)。
     * customWidth の列がないシートでは広げない。試験 fb3-b9〜b18 (Fableレポート3 の 9。夜間 n0426・n0627・n1295)。
     */
    private int[] fitPx;
    static final boolean VALUE_FIT = !"0".equals(System.getProperty("x2p.valuefit"));
    static final boolean VALUE_FIT_FORMULA = "1".equals(System.getProperty("x2p.valuefitf"));
    /**
     * 広げる幅 = 文字の幅 + セルの大きさの標準フォントの数字幅 × 2 (+ セルの大きさが標準フォントと同じなら列の余白 2 × (mdw / 4) + 1)。
     * 試験 fb3-b18: 足す量は 8pt で 12px、11pt (標準と同じ) で 21px、14pt で 19px。10 通り中 9 通りが 1px 以内 (Calibri 11pt だけ 2px 広い)。
     */
    static final double VALUE_FIT_DIGITS = Double.parseDouble(System.getProperty("x2p.valuefitd", "2"));

    private void valueFit(org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet ws) {
        boolean[] cw = new boolean[maxCol], elem = new boolean[maxCol];
        boolean any = false;
        for (CTCols cols : ws.getColsArray())
            for (CTCol c : cols.getColArray())
                for (long i = c.getMin() - 1; i < Math.min(c.getMax(), maxCol); i++) if (c.isSetWidth()) elem[(int) i] = true;
        for (CTCols cols : ws.getColsArray())
            for (CTCol c : cols.getColArray()) {
                if (!(c.isSetCustomWidth() && c.getCustomWidth())) continue;
                any = true;
                if (c.getMin() <= 1 && c.getMax() >= 16384) continue;   // 全列の指定は、列そのものは customWidth のない列として扱う (試験 fb3-b16 i_max_cw は広がる)
                for (long i = c.getMin() - 1; i < Math.min(c.getMax(), maxCol); i++) cw[(int) i] = true;
            }
        if (!any) return;
        FontManager fm = model.fontManager();
        int[] out = null;
        for (var row : sheet) {
            for (Cell cell : row) {
                int c = cell.getColumnIndex();
                if (c >= maxCol || cw[c] || colChars[c] <= 0 || mergedAt(cell.getRowIndex(), c) != null) continue;
                // 数式のセルは広げない (夜間 n3147・n0600: 数式の結果のエラー値で Excel は広げない)
                if (cell.getCellType() == CellType.FORMULA && !VALUE_FIT_FORMULA) continue;
                CellType t = cell.getCellType() == CellType.FORMULA ? cell.getCachedFormulaResultType() : cell.getCellType();
                String text;
                boolean num = false;
                XSSFCellStyle st = (XSSFCellStyle) cell.getCellStyle();
                if (st != null && (st.getShrinkToFit() || st.getWrapText() || st.getRotation() != 0)) continue;
                if (t == CellType.BOOLEAN) {
                    // セルのチェックボックス (書式の extLst の xfpb:xfComplement) は文字を描かないので広げない (夜間 n3207: Excel で保存した同じ様式)
                    var xf = st != null ? st.getCoreXf() : null;
                    if (xf != null && xf.isSetExtLst() && xf.getExtLst().xmlText().contains("xfComplement")) continue;
                    text = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                }
                else if (t == CellType.ERROR) {
                    try {
                        text = org.apache.poi.ss.usermodel.FormulaError.forInt(cell.getErrorCellValue()).getString();
                    } catch (RuntimeException e) {
                        continue;
                    }
                } else if (t == CellType.NUMERIC && st != null && st.getDataFormat() == 0 && elem[c]) {
                    // 数値は col 要素 (幅の指定) のある列だけ。col 要素のない列は長い数値の規則 (numWiden、試験 exp25) に任せる
                    double v = cell.getNumericCellValue();
                    if (v != Math.rint(v) || Math.abs(v) >= 1e11) continue;   // 小数は桁を減らして表示する (広げない。fb3-b18 k_f2)
                    text = String.valueOf((long) v);
                    num = true;
                } else continue;
                XSSFFont f = st != null ? st.getFont() : null;
                if (f == null) continue;
                double size = f.getFontHeight() / 20.0;
                String face = num ? model.defaultFont() : metricName(f);   // 数値は標準フォントの数字で測る (fb3-b18 k_n5_*: 書体によらない)
                float pt = (float) (screenPpem(face, size) * 0.75);
                FontManager.Spec sp = fm.spec(face, !num && f.getBold(), !num && f.getItalic());
                int w = 0;
                for (int i = 0; i < text.length(); i++) w += Math.round(sp.charWidth(text.charAt(i), pt, fm) / 0.75f);
                double colPx = Math.floor(((256 * colChars[c] + Math.floor(128.0 / mdw)) / 256) * mdw);
                int pad = 2 * (mdw / 4) + 1;
                if (w <= colPx - pad) continue;
                int d = Math.round(fm.spec(model.defaultFont(), false, false).charWidth('0', pt, fm) / 0.75f);
                boolean stdSize = Math.abs(size - model.defaultSize()) < 0.01;
                int need = (int) Math.round(w + VALUE_FIT_DIGITS * d + (stdSize ? pad : 0));
                if (out == null) out = new int[maxCol];
                out[c] = Math.max(out[c], need);
            }
        }
        fitPx = out;
    }

    /** 行に、文字列のある折り返しのセルがあるか (結合のセルは数えない)。 */
    /**
     * 折り返しの「選択範囲内で中央」のセル (文字のあるもの) があるか。保存値が自動調整の 1.5 倍未満 (1 行に収まる) のときだけ
     * 測り直す (実ファイル munp_yousiki1-4-2)。それより高い行は複数行の保存値のまま (総当たり sw-merge の 131.25pt)。
     */
    private boolean hasSelWrap(XSSFRow row) {
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.STRING || cell.getStringCellValue().isEmpty()) continue;
            XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
            if (cs.getWrapText() && cs.getRotation() == 0 && cs.getAlignment() == org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER_SELECTION
                    && mergedAt(cell.getRowIndex(), cell.getColumnIndex()) == null) return true;
        }
        return false;
    }

    private boolean hasWrapText(XSSFRow row) {
        for (Cell cell : row) {
            if (cell.getCellType() != CellType.STRING || cell.getStringCellValue().isEmpty()) continue;
            XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
            if (ROT_WRAP_KEEP && cs.getRotation() != 0) continue;
            // 選択範囲内で中央のセルも数え直しに数えない (rewrapPx と同じ。総当たり sw-merge: Excel の保存値 131.25pt のまま)
            if (SEL_WRAP_KEEP && cs.getAlignment() == org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER_SELECTION) continue;
            if (cs.getWrapText() && mergedAt(cell.getRowIndex(), cell.getColumnIndex()) == null) return true;
        }
        return false;
    }

    /** セルのない行が列の書式 (塗りつぶし・罫線) を継ぐとき、その行までを使用範囲に数える (0 = 無効、1 = row 要素のある行、2 = customHeight の行だけ)。 */
    static final int COL_STYLE_ROWS = Integer.getInteger("x2p.colstylerows", 2);

    /** 書式に罫線か塗りつぶしがあるか。 */
    static boolean styleSignificant(org.apache.poi.ss.usermodel.CellStyle cs) {
        if (!(cs instanceof XSSFCellStyle s)) return false;
        if (s.getBorderTop() != BorderStyle.NONE || s.getBorderBottom() != BorderStyle.NONE
                || s.getBorderLeft() != BorderStyle.NONE || s.getBorderRight() != BorderStyle.NONE) return true;
        return s.getFillPattern() != FillPatternType.NO_FILL;
    }

    /** 使用範囲に数えるセルか (値・数式、または罫線・塗りつぶしがある)。 */
    static boolean significant(Cell cell) {
        CellType t = cell.getCellType();
        // 数式は結果が "" でも Excel の使用範囲に数える
        if (t == CellType.FORMULA) return true;
        if (t != CellType.BLANK && !(t == CellType.STRING && cell.getStringCellValue().isEmpty())) return true;
        XSSFCellStyle s = (XSSFCellStyle) cell.getCellStyle();
        if (s == null) return false;
        if (s.getBorderTop() != BorderStyle.NONE || s.getBorderBottom() != BorderStyle.NONE
                || s.getBorderLeft() != BorderStyle.NONE || s.getBorderRight() != BorderStyle.NONE) return true;
        return s.getFillPattern() != FillPatternType.NO_FILL || CellPainter.hasGradient(cell);
    }
}
