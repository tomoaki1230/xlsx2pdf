package xlsx2pdf.render;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBreak;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageMargins;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPrintOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

/** Excel が印刷するときのシートのページ設定。長さはすべて pt。 */
public final class PrintSetup {
    /** 用紙の幅・高さ (向きを反映済み)。 */
    public double paperW = 595.28, paperH = 841.89;
    /** 余白 (左・右・上・下・ヘッダー・フッター)。 */
    public double left = 0.7 * 72, right = 0.7 * 72, top = 0.75 * 72, bottom = 0.75 * 72, header = 0.3 * 72, footer = 0.3 * 72;
    /** 拡大縮小率 (%)。 */
    public int scale = 100;
    /** pageSetup の scale 属性があるか (ページに合わせる設定では、保存時に Excel が計算した倍率)。 */
    public boolean scaleSet;
    /** 「次のページ数に合わせて印刷」か。 */
    public boolean fitToPage;
    /** 合わせるページ数 (横・縦。0 は自動)。 */
    public int fitW = 1, fitH = 1;
    /** 水平・垂直に中央揃え。 */
    public boolean centerH, centerV;
    /** 枠線・行列番号を印刷する。 */
    public boolean gridlines, headings;
    /** 簡易印刷 (pageSetup の draft)。Excel はセルの文字だけを印刷する (塗り・罫線・枠線・図形・グラフを印刷しない)。 */
    public boolean draft;
    /** 右から左のシート (sheetView の rightToLeft)。列を右から左へ並べる。 */
    public boolean rtl;
    /** ページの順序が「左から右」。 */
    public boolean overThenDown;
    /** 先頭ページ番号 (null = 自動)。 */
    public Integer firstPageNumber;
    /** セルのエラーの印刷方法 (displayed, blank, dash, NA)。 */
    public String errors = "displayed";
    /** 白黒印刷。 */
    public boolean blackAndWhite;
    /** ゼロ値を表示する。 */
    public boolean showZeros = true;
    /** 手動改ページ (この行の前で改ページ、0 始まり)。 */
    public final List<Integer> rowBreaks = new ArrayList<>();
    /** 手動改ページ (この列の前で改ページ、0 始まり)。 */
    public final List<Integer> colBreaks = new ArrayList<>();
    /** 印刷範囲 (なければ空)。 */
    public final List<CellRangeAddress> printAreas = new ArrayList<>();
    /** 印刷タイトルの行・列 (なければ null)。 */
    public CellRangeAddress titleRows, titleCols;
    /** コメントの印刷方法 (none, asDisplayed, atEnd)。 */
    public String comments = "none";
    /** レイアウトに使う解像度 (dpi)。 */
    public int hdpi = 600, vdpi = 600;
    /** ファイルの pageSetup の horizontalDpi/verticalDpi を使う (既定は常に 600 dpi)。 */
    public static boolean useFileDpi = false;

    private static final double MM = 72 / 25.4;
    /**
     * Excel の PDF 出力は、内容を用紙の上端・左端から印刷可能領域 (約 0.16 インチ) より近くに置かない。
     * ページ分割には設定どおりの余白を使う。
     */
    public static final double MIN_MARGIN = 0.16 * 72;
    /**
     * 余白の下限 (pt)。--author-env では {@link #MIN_MARGIN}、既定では 0。
     */
    public double minMargin = MIN_MARGIN;

    /** SpreadsheetML の paperSize コードに対応する用紙サイズ {幅, 高さ} (縦向き、pt)。不明なら A4。 */
    static double[] paper(int code) {
        return switch (code) {
            case 1, 2, 18 -> new double[] {612, 792};
            case 3, 17 -> new double[] {792, 1224};
            case 4 -> new double[] {1224, 792};
            case 5 -> new double[] {612, 1008};
            case 6 -> new double[] {396, 612};
            case 7 -> new double[] {522, 756};
            case 8 -> new double[] {297 * MM, 420 * MM};
            case 9, 10 -> new double[] {210 * MM, 297 * MM};
            case 11 -> new double[] {148 * MM, 210 * MM};
            case 12 -> new double[] {257 * MM, 364 * MM};
            case 13 -> new double[] {182 * MM, 257 * MM};
            case 14 -> new double[] {612, 936};
            case 15 -> new double[] {215 * MM, 275 * MM};
            case 16 -> new double[] {720, 1008};
            case 19 -> new double[] {279, 639};
            case 20 -> new double[] {297, 684};
            case 27 -> new double[] {110 * MM, 220 * MM};
            case 28 -> new double[] {162 * MM, 229 * MM};
            case 34 -> new double[] {176 * MM, 250 * MM};
            case 37 -> new double[] {279, 540};
            case 43 -> new double[] {100 * MM, 148 * MM};
            case 63 -> new double[] {297 * MM, 420 * MM};
            case 66 -> new double[] {420 * MM, 594 * MM};
            case 70 -> new double[] {105 * MM, 148 * MM};
            case 41 -> new double[] {1224, 792};
            default -> new double[] {210 * MM, 297 * MM};
        };
    }

    /**
     * シートのページ設定を読む。
     *
     * @param sheetIndex   シートの番号 (印刷範囲・印刷タイトルの名前の照合に使う)
     * @param defaultPaper ファイルに用紙サイズがないときの用紙コード
     */
    public static PrintSetup read(XSSFWorkbook wb, XSSFSheet sheet, int sheetIndex, int defaultPaper) {
        return read(wb, sheet, sheetIndex, defaultPaper, 0);
    }

    /**
     * シートのページ設定を、出力環境のプリンタ解像度を指定して読む。
     *
     * @param printerDpi 出力に使ったプリンタの解像度 (Microsoft Print to PDF なら 600、多くのレーザープリンタは 300 か 600)。
     *                   0 なら既定
     */
    public static PrintSetup read(XSSFWorkbook wb, XSSFSheet sheet, int sheetIndex, int defaultPaper, int printerDpi) {
        PrintSetup p = new PrintSetup();
        CTWorksheet ws = sheet.getCTWorksheet();
        CTPageSetup ps = ws.isSetPageSetup() ? ws.getPageSetup() : null;
        int paper = defaultPaper;
        boolean landscape = false;
        if (ps != null) {
            if (ps.isSetPaperSize() && ps.getPaperSize() > 0) paper = (int) ps.getPaperSize();
            if (ps.isSetOrientation()) landscape = "landscape".equals(ps.getOrientation().toString());
            if (ps.isSetScale()) { p.scale = (int) ps.getScale(); p.scaleSet = true; }
            if (ps.isSetFitToWidth()) p.fitW = (int) ps.getFitToWidth();
            if (ps.isSetFitToHeight()) p.fitH = (int) ps.getFitToHeight();
            if (ps.isSetPageOrder()) p.overThenDown = "overThenDown".equals(ps.getPageOrder().toString());
            if (ps.isSetUseFirstPageNumber() && ps.getUseFirstPageNumber() && ps.isSetFirstPageNumber())
                p.firstPageNumber = (int) ps.getFirstPageNumber();
            if (ps.isSetErrors()) p.errors = ps.getErrors().toString();
            if (ps.isSetBlackAndWhite()) p.blackAndWhite = ps.getBlackAndWhite();
            if (ps.isSetDraft()) p.draft = ps.getDraft();
            if (ps.isSetCellComments()) p.comments = ps.getCellComments().toString();
            if (ps.isSetHorizontalDpi() && validDpi(ps.getHorizontalDpi())) p.hdpi = (int) ps.getHorizontalDpi();
            if (ps.isSetVerticalDpi() && validDpi(ps.getVerticalDpi())) p.vdpi = (int) ps.getVerticalDpi();
            // Excel の PDF 出力は、ファイルに別の解像度が書かれていても多くは 600 dpi でレイアウトする
            // (そのプリンタが出力した PC にないことが多い)。--use-file-dpi のときだけファイルの値を使う
            if (!useFileDpi) { p.hdpi = 600; p.vdpi = 600; }
            if (ps.isSetPaperWidth() && ps.isSetPaperHeight()) {
                double w = parseLen(ps.getPaperWidth()), h = parseLen(ps.getPaperHeight());
                if (w > 0 && h > 0) { p.paperW = w; p.paperH = h; paper = -1; }
            }
        }
        if (paper != -1) {
            double[] d = paper(paper);
            p.paperW = d[0];
            p.paperH = d[1];
        }
        if (landscape) { double t = p.paperW; p.paperW = p.paperH; p.paperH = t; }
        if (p.scale < 10 || p.scale > 400) p.scale = 100;
        if (ws.isSetSheetPr() && ws.getSheetPr().isSetPageSetUpPr() && ws.getSheetPr().getPageSetUpPr().isSetFitToPage())
            p.fitToPage = ws.getSheetPr().getPageSetUpPr().getFitToPage();
        if (ws.isSetPageMargins()) {
            CTPageMargins m = ws.getPageMargins();
            p.left = m.getLeft() * 72; p.right = m.getRight() * 72; p.top = m.getTop() * 72; p.bottom = m.getBottom() * 72;
            p.header = m.getHeader() * 72; p.footer = m.getFooter() * 72;
        }
        if (ws.isSetPrintOptions()) {
            CTPrintOptions o = ws.getPrintOptions();
            p.centerH = o.isSetHorizontalCentered() && o.getHorizontalCentered();
            p.centerV = o.isSetVerticalCentered() && o.getVerticalCentered();
            p.gridlines = o.isSetGridLines() && o.getGridLines() && (!o.isSetGridLinesSet() || o.getGridLinesSet());
            p.headings = o.isSetHeadings() && o.getHeadings();
        }
        if (ws.isSetSheetViews() && ws.getSheetViews().sizeOfSheetViewArray() > 0) {
            CTSheetView v = ws.getSheetViews().getSheetViewArray(0);
            if (v.isSetShowZeros()) p.showZeros = v.getShowZeros();
            if (v.isSetRightToLeft()) p.rtl = v.getRightToLeft();
        }
        if (ws.isSetRowBreaks()) for (CTBreak b : ws.getRowBreaks().getBrkArray()) if (!b.isSetMan() || b.getMan()) p.rowBreaks.add((int) b.getId());
        if (ws.isSetColBreaks()) for (CTBreak b : ws.getColBreaks().getBrkArray()) if (!b.isSetMan() || b.getMan()) p.colBreaks.add((int) b.getId());
        // 名前定義 (印刷範囲・印刷タイトル)
        for (XSSFName n : wb.getAllNames()) {
            if (n.getSheetIndex() != sheetIndex) continue;
            String f = n.getRefersToFormula();
            if (f == null) continue;
            if ("_xlnm.Print_Area".equalsIgnoreCase(n.getNameName())) {
                // 壊れた参照 (#REF!) を含む印刷範囲は、Excel は全体を無効として使用範囲を印刷する (A1:U72,#REF!)
                if (f.contains("#REF!")) continue;
                if (f.replaceAll("'[^']*'", "").contains("(")) {   // シート名 ('3-9(1)' など) の括弧は数えない
                    // 数式の印刷範囲 (OFFSET(…,COUNTA(…)) など): Excel は評価した範囲を印刷する
                    CellRangeAddress a = evalArea(wb, sheet, f);
                    if (a != null) p.printAreas.add(a);
                    continue;
                }
                for (String part : splitRefs(f)) {
                    CellRangeAddress a = parseRange(part);
                    if (a != null && a.getFirstRow() >= 0 && a.getFirstColumn() >= 0) p.printAreas.add(a);
                }
            } else if ("_xlnm.Print_Titles".equalsIgnoreCase(n.getNameName())) {
                for (String part : splitRefs(f)) {
                    String r = part.contains("!") ? part.substring(part.lastIndexOf('!') + 1) : part;
                    r = r.replace("$", "");
                    String[] ab = r.split(":");
                    if (ab.length != 2) continue;
                    if (ab[0].matches("\\d+")) {
                        p.titleRows = new CellRangeAddress(Integer.parseInt(ab[0]) - 1, Integer.parseInt(ab[1]) - 1, -1, -1);
                    } else if (ab[0].matches("[A-Za-z]+")) {
                        p.titleCols = new CellRangeAddress(-1, -1, colIndex(ab[0]), colIndex(ab[1]));
                    } else {
                        CellRangeAddress a = parseRange(part);
                        if (a != null) {
                            if (a.getFirstColumn() == 0 && a.getLastColumn() >= 16383)
                                p.titleRows = new CellRangeAddress(a.getFirstRow(), a.getLastRow(), -1, -1);
                            else if (a.getFirstRow() == 0 && a.getLastRow() >= 1048575)
                                p.titleCols = new CellRangeAddress(-1, -1, a.getFirstColumn(), a.getLastColumn());
                        }
                    }
                }
            }
        }
        if (printerDpi > 0) { p.hdpi = printerDpi; p.vdpi = printerDpi; }
        return p;
    }

    /**
     * 有効な解像度か。DEVMODE の解像度は負の品質コード (-1 ドラフト〜-4 高品質) のことがあり、4294967292 などとして
     * 保存される。それらは「プリンタの既定」(600 dpi) を意味する。
     */
    static boolean validDpi(long dpi) {
        return dpi >= 72 && dpi <= 4800;
    }

    /** 用紙名 (A4, A3, B4, B5, Letter, Legal など、または SpreadsheetML の番号) から用紙コードを返す。 */
    public static int paperCode(String name) {
        return switch (name.trim().toUpperCase(java.util.Locale.ROOT)) {
            case "LETTER" -> 1;
            case "LEGAL" -> 5;
            case "A3" -> 8;
            case "A4" -> 9;
            case "A5" -> 11;
            case "B4" -> 12;
            case "B5" -> 13;
            case "TABLOID" -> 3;
            default -> Integer.parseInt(name.trim());
        };
    }

    /** 単位付きの長さ (mm, cm, in。単位なしは pt) を pt にする。解釈できなければ -1。 */
    private static double parseLen(String s) {
        try {
            if (s.endsWith("mm")) return Double.parseDouble(s.substring(0, s.length() - 2)) * MM;
            if (s.endsWith("cm")) return Double.parseDouble(s.substring(0, s.length() - 2)) * MM * 10;
            if (s.endsWith("in")) return Double.parseDouble(s.substring(0, s.length() - 2)) * 72;
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    /** 列名 (A, B, ..., AA) を 0 始まりの列番号にする。 */
    static int colIndex(String letters) {
        int n = 0;
        for (char c : letters.toUpperCase().toCharArray()) n = n * 26 + (c - 'A' + 1);
        return n - 1;
    }

    /** 数式の範囲 (名前定義の OFFSET など) を評価してセル範囲にする。評価できなければ null (使用範囲を印刷する)。 */
    static CellRangeAddress evalArea(XSSFWorkbook wb, XSSFSheet sheet, String f) {
        try {
            var fe = wb.getCreationHelper().createFormulaEvaluator();
            var v = ((org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator) fe)._getWorkbookEvaluator().evaluateList(f,
                    new org.apache.poi.ss.util.CellReference(sheet.getSheetName(), 0, 0, false, false), new CellRangeAddress(0, 0, 0, 0));
            if (v instanceof org.apache.poi.ss.formula.eval.AreaEval a && a.getFirstRow() >= 0 && a.getFirstColumn() >= 0)
                return new CellRangeAddress(a.getFirstRow(), a.getLastRow(), a.getFirstColumn(), a.getLastColumn());
            if (v instanceof org.apache.poi.ss.formula.eval.RefEval r && r.getRow() >= 0 && r.getColumn() >= 0)
                return new CellRangeAddress(r.getRow(), r.getRow(), r.getColumn(), r.getColumn());
        } catch (RuntimeException e) {
            // 評価できない数式は印刷範囲なしとする
        }
        return null;
    }

    /** カンマ区切りの参照を分割する (引用符で囲まれたシート名の中のカンマは区切りとしない)。 */
    static List<String> splitRefs(String f) {
        List<String> out = new ArrayList<>();
        StringBuilder cur = new StringBuilder();
        boolean q = false;
        for (char c : f.toCharArray()) {
            if (c == '\'') q = !q;
            if (c == ',' && !q) { out.add(cur.toString()); cur.setLength(0); continue; }
            cur.append(c);
        }
        if (cur.length() > 0) out.add(cur.toString());
        return out;
    }

    /**
     * 参照 (シート名付き可、$ 可) をセル範囲にする。行全体 (1:3) と列全体 (A:C) にも対応する。
     *
     * @return セル範囲。#REF や解釈できない参照なら null
     */
    static CellRangeAddress parseRange(String part) {
        String r = part.contains("!") ? part.substring(part.lastIndexOf('!') + 1) : part;
        r = r.replace("$", "").trim();
        if (r.startsWith("#REF")) return null;
        try {
            if (!r.contains(":")) r = r + ":" + r;
            String[] ab = r.split(":");
            if (ab[0].matches("\\d+") && ab[1].matches("\\d+"))
                return new CellRangeAddress(Integer.parseInt(ab[0]) - 1, Integer.parseInt(ab[1]) - 1, 0, 16383);
            if (ab[0].matches("[A-Za-z]+") && ab[1].matches("[A-Za-z]+"))
                return new CellRangeAddress(0, 1048575, colIndex(ab[0]), colIndex(ab[1]));
            return CellRangeAddress.valueOf(r);
        } catch (RuntimeException e) {
            return null;
        }
    }
}
