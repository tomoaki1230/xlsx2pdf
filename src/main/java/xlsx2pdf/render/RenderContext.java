package xlsx2pdf.render;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** ブック全体で共有する描画用の状態 (フォント、テーマ、標準フォント、印刷モデルなど)。 */
public final class RenderContext {
    /** 変換元のブック。 */
    public final XSSFWorkbook wb;
    /** 出力先の PDF。 */
    public final PDDocument doc;
    /** フォントの検索・埋め込み。 */
    public final FontManager fonts;
    /** テーマの色とフォント。 */
    public final Theme theme;
    /** セル値の表示形式。 */
    public final ValueFormatter formatter;
    /** 標準スタイルのフォント名。 */
    public final String defaultFontName;
    /** 標準スタイルのフォントサイズ (pt)。 */
    public final double defaultFontSize;
    /** 標準フォントの最大数字幅 (96 dpi の画面ピクセル)。 */
    public final int mdw;
    /** Excel の印刷時の列幅・行高の換算モデル。 */
    public final PrinterModel printer;
    /** 日本語のブック (標準フォントまたはテーマの東アジアフォントが日本語) か。 */
    public final boolean japanese;
    /** 埋め込み済みの画像 (同じ画像を 1 回だけ埋め込むためのキャッシュ)。 */
    public final Map<Object, PDImageXObject> images = new HashMap<>();
    /** 変換元のファイル。 */
    public final File source;
    /** ブック名 (ヘッダーの &amp;F)。ストリームから変換するときは Options の値。なければ null。 */
    public String sourceName;
    /** 時間の上限を確かめる処理 (上限を超えていれば LimitExceededException)。 */
    public Runnable deadline = () -> { };

    /** 時間の上限を超えていれば変換を中止する (セルの行・図形ごとに呼ぶ)。 */
    public void checkDeadline() {
        deadline.run();
    }

    /**
     * ブックを最後に保存した Excel のメジャーバージョンを返す (docProps/app.xml の AppVersion)。
     *
     * @return 14 = 2010、15 = 2013、16 = 2016 以降、0 = 不明
     */
    /** Excel 以外で保存したブックの手で決めていない行の高さを測り直す (exp111・113・115)。 */
    static final boolean NON_EXCEL_REFIT = true;

    static int appVersion(XSSFWorkbook wb) {
        try {
            String v = wb.getProperties().getExtendedProperties().getUnderlyingProperties().getAppVersion();
            if (v != null) return Integer.parseInt(v.substring(0, v.indexOf('.') > 0 ? v.indexOf('.') : v.length()));
        } catch (RuntimeException ignored) {
            // 不明
        }
        return 0;
    }

    /** Excel 以外のアプリ (LibreOffice、JUST Calc、Apache POI など) で保存したブックか (docProps/app.xml の Application。なければ不明で false)。 */
    /** 名前のないフォント (<font></font>) を、日本語版の Excel と同じく ＭＳ Ｐゴシック 11pt とみなす。 */
    static final boolean EMPTY_FONT_MSPG = true;

    /**
     * 名前の要素のないフォントに ＭＳ Ｐゴシック (大きさもなければ 11pt) を入れる (読み込んだブックの中だけ)。
     * 日本語版の Excel は、標準の書式のフォントが空のブックを標準の高さ 13.5pt・幅 8.38 (ＭＳ Ｐゴシック 11) で開く (夜間 n0248・n0566・n0599・n0871)。
     */
    static final int CELL_S0 = 1;

    /**
     * s 属性のないセルを書式 0 にする。POI は s のないセルの書式を列の書式で補うが、Excel は書式 0 を使う
     * (夜間 n0131: 列の書式が灰色の塗りで、s のないセルを Excel は白で印刷する。値のある文字のセルも同じ)。
     * ただし n1799 では値のある s のないセルを Excel が列の書式のフォントで印刷しており、規則は未確定。列の書式のあるシートだけ調べる。
     */
    private static void explicitDefaultStyle(XSSFWorkbook wb) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            var ws = wb.getSheetAt(i).getCTWorksheet();
            boolean styledCols = false;
            for (var cols : ws.getColsArray())
                for (var c : cols.getColArray()) if (c.isSetStyle() && c.getStyle() != 0) styledCols = true;
            if (!styledCols || ws.getSheetData() == null) continue;
            for (var row : ws.getSheetData().getRowArray())
                for (var c : row.getCArray()) if (!c.isSetS()) c.setS(0);
        }
    }

    /**
     * Windows にない書体名を、Excel (GDI のフォントマッパー) と同じく charset と family で決まる書体に置き換える (Fableレポート A。
     * 試験 fb-missfont 196 件)。テーマの書体 (scheme) は対象外。以後の寸法の表と描画はすべて置き換え後の名前で動く。
     */
    static final boolean FULLNAME_RULE = !"0".equals(System.getProperty("x2p.fullname"));

    static void substituteMissingFonts(XSSFWorkbook wb) {
        try {
            for (XSSFFont f : wb.getStylesSource().getFonts()) {
                var ct = f.getCTFont();
                if (ct.sizeOfNameArray() == 0 || ct.sizeOfSchemeArray() > 0) continue;
                String name = ct.getNameArray(0).getVal();
                if (name == null || name.isBlank()) continue;
                int cs = ct.sizeOfCharsetArray() > 0 ? (int) ct.getCharsetArray(0).getVal() : 0;
                int fam = ct.sizeOfFamilyArray() > 0 ? (int) ct.getFamilyArray(0).getVal() : 0;
                boolean inst = FontManager.installed(name);
                // スタイル付きの名前は、charset がなく family があると Excel は見つけられない
                if (inst && FULLNAME_RULE && ct.sizeOfCharsetArray() == 0 && fam != 0 && FontManager.fullNameOnly(name)) inst = false;
                if (inst) continue;
                String sub;
                if (cs == 128 || cs == 1) sub = fam == 1 ? "ＭＳ Ｐ明朝" : "ＭＳ ゴシック";
                else if (cs == 0 || cs == 161 || cs == 162 || cs == 186 || cs == 204 || cs == 238)
                    sub = fam == 1 ? "Times New Roman" : fam == 3 ? "Courier New" : "Arial";
                else continue;
                ct.getNameArray(0).setVal(sub);
            }
        } catch (RuntimeException e) {
            // 読めないときはそのまま
        }
    }

    /**
     * グラフシートの余白・ページ設定・ヘッダー/フッター (画像を含む) を、POI がグラフシートに持たせる空のワークシートへ写す
     * (印刷設定とヘッダーはワークシートと同じ処理で読むため。exp122 chartsheet: 横向きとヘッダーが出ていなかった)。
     * ページ設定の型 (CTCsPageSetup) は POI の軽量版の部品にないので、XML の属性を読む。
     */
    private static void syncChartSheets(XSSFWorkbook wb) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (!(wb.getSheetAt(i) instanceof org.apache.poi.xssf.usermodel.XSSFChartSheet cs)) continue;
            try {
                var cc = cs.getCTChartsheet();
                var ws = cs.getCTWorksheet();
                if (cc.isSetHeaderFooter() && !ws.isSetHeaderFooter()) ws.setHeaderFooter(cc.getHeaderFooter());
                if (cc.isSetLegacyDrawingHF() && !ws.isSetLegacyDrawingHF()) ws.setLegacyDrawingHF(cc.getLegacyDrawingHF());
                // グラフシートの向きの既定は横 (pageSetup がないブック。rand30 rd08・rd27、exp124 cs_default)
                var ps0 = ws.isSetPageSetup() ? ws.getPageSetup() : ws.addNewPageSetup();
                if (!ps0.isSetOrientation()) ps0.setOrientation(org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation.LANDSCAPE);
                for (org.w3c.dom.Node n = cc.getDomNode().getFirstChild(); n != null; n = n.getNextSibling()) {
                    if (!(n instanceof org.w3c.dom.Element e)) continue;
                    switch (e.getLocalName()) {
                        case "pageMargins" -> {
                            var m = ws.isSetPageMargins() ? ws.getPageMargins() : ws.addNewPageMargins();
                            m.setLeft(num(e, "left", 0.7)); m.setRight(num(e, "right", 0.7)); m.setTop(num(e, "top", 0.75));
                            m.setBottom(num(e, "bottom", 0.75)); m.setHeader(num(e, "header", 0.3)); m.setFooter(num(e, "footer", 0.3));
                        }
                        case "pageSetup" -> {
                            var ps = ws.isSetPageSetup() ? ws.getPageSetup() : ws.addNewPageSetup();
                            if (e.hasAttribute("paperSize")) ps.setPaperSize((long) num(e, "paperSize", 9));
                            if (e.hasAttribute("orientation"))
                                ps.setOrientation(org.openxmlformats.schemas.spreadsheetml.x2006.main.STOrientation.Enum.forString(e.getAttribute("orientation")));
                            if (e.hasAttribute("firstPageNumber")) ps.setFirstPageNumber((long) num(e, "firstPageNumber", 1));
                            if (e.hasAttribute("useFirstPageNumber")) ps.setUseFirstPageNumber(bool(e, "useFirstPageNumber"));
                            if (e.hasAttribute("blackAndWhite")) ps.setBlackAndWhite(bool(e, "blackAndWhite"));
                            if (e.hasAttribute("draft")) ps.setDraft(bool(e, "draft"));
                        }
                        default -> { }
                    }
                }
            } catch (RuntimeException e) {
                // 写せなければ従来どおり (縦向き・既定の余白)
            }
        }
    }

    private static double num(org.w3c.dom.Element e, String a, double d) {
        try {
            return e.hasAttribute(a) ? Double.parseDouble(e.getAttribute(a)) : d;
        } catch (NumberFormatException x) {
            return d;
        }
    }

    private static boolean bool(org.w3c.dom.Element e, String a) {
        String v = e.getAttribute(a);
        return "1".equals(v) || "true".equals(v);
    }

    private static void fillEmptyFonts(XSSFWorkbook wb) {
        try {
            for (XSSFFont f : wb.getStylesSource().getFonts()) {
                var ct = f.getCTFont();
                if (ct.sizeOfNameArray() > 0) continue;
                ct.addNewName().setVal("ＭＳ Ｐゴシック");
                if (ct.sizeOfSzArray() == 0) ct.addNewSz().setVal(11);
            }
        } catch (RuntimeException e) {
            // 読めないときはそのまま
        }
    }

    /** 名前が Microsoft でもテーマの部品がないブックを Excel 以外で保存したブックとみなす。 */
    static final boolean NO_THEME_NON_EXCEL = true;

    static boolean nonExcel(XSSFWorkbook wb) {
        try {
            String a = wb.getProperties().getExtendedProperties().getUnderlyingProperties().getApplication();
            if (a != null && !a.isBlank() && !a.startsWith("Microsoft")) return true;
            // 名前は Microsoft でもテーマの部品がないブックは、Excel 以外が作ったもの (Excel は必ずテーマを書く。夜間 n0248 など)
            return NO_THEME_NON_EXCEL && a != null && a.startsWith("Microsoft")
                    && wb.getPackage().getPartsByContentType("application/vnd.openxmlformats-officedocument.theme+xml").isEmpty();
        } catch (RuntimeException e) {
            return false;
        }
    }

    /**
     * 描画の状態を作る。
     *
     * @param wb       変換元のブック
     * @param doc      出力先の PDF
     * @param fontDirs 追加で検索するフォントフォルダ
     * @param source   変換元のファイル
     */
    /** 標準のスタイル (builtinId 0) の名前が「標準」か。 */
    private static boolean normalStyleJa(XSSFWorkbook wb) {
        try {
            var css = wb.getStylesSource().getCTStylesheet().getCellStyles();
            if (css == null) return false;
            for (var cs : css.getCellStyleArray())
                if (cs.isSetBuiltinId() && cs.getBuiltinId() == 0) return "標準".equals(cs.getName());
        } catch (RuntimeException e) {
            return false;
        }
        return false;
    }

    /**
     * セルのフォントの書体名。テーマのフォント (scheme minor / major) でテーマに日本語 (Jpan) の書体があれば、日本語版の Excel は
     * 英数字も含めてその書体で描く (exp100: 名前が Calibri でも ＭＳ Ｐゴシック。テーマの ea の欄や標準のスタイル名にはよらない)。
     */
    public String cellFontName(org.apache.poi.xssf.usermodel.XSSFFont f) {
        org.apache.poi.ss.usermodel.FontScheme s = f.getScheme();
        if (s == org.apache.poi.ss.usermodel.FontScheme.MINOR && printer.minorJpan != null) return printer.minorJpan;
        if (s == org.apache.poi.ss.usermodel.FontScheme.MAJOR && printer.majorJpan != null) return printer.majorJpan;
        return f.getFontName();
    }

    /** 標準フォントを「標準」スタイル (cellStyles の builtinId 0 → cellStyleXfs) から取る。無効なら cellXfs[0] のフォント。 */
    static final boolean NORMAL_STYLE_FONT = !"0".equals(System.getProperty("x2p.normalfont"));

    /**
     * 「標準」スタイルのフォント。Excel は列幅の単位 (数字幅) と標準行を、cellXfs[0] ではなく「標準」スタイルのフォントで決める
     * (夜間 n0334: cellXfs[0] は scheme のない Calibri、標準スタイルは scheme minor の Calibri → 日本語版では ＭＳ Ｐゴシックで数字幅 8)。
     */
    private static XSSFFont normalStyleFont(XSSFWorkbook wb) {
        try {
            var st = wb.getStylesSource();
            var ss = st.getCTStylesheet();
            long xfId = 0;
            boolean found = false;
            if (ss.isSetCellStyles()) {
                for (var cs : ss.getCellStyles().getCellStyleArray()) {
                    if (cs.isSetBuiltinId() && cs.getBuiltinId() == 0) { xfId = cs.getXfId(); found = true; break; }
                }
            }
            if (!ss.isSetCellStyleXfs() || xfId < 0 || xfId >= ss.getCellStyleXfs().sizeOfXfArray()) return null;
            if (!found && ss.getCellStyleXfs().sizeOfXfArray() == 0) return null;
            var xf = ss.getCellStyleXfs().getXfArray((int) xfId);
            int fid = xf.isSetFontId() ? (int) xf.getFontId() : 0;
            if (fid < 0 || fid >= st.getFonts().size()) return null;
            return st.getFontAt(fid);
        } catch (RuntimeException e) {
            return null;
        }
    }

    /** 数式を再計算した (--recalc)。グラフはキャッシュではなく参照先のセルの値を使う。 */
    public boolean recalc;

    public RenderContext(XSSFWorkbook wb, PDDocument doc, List<File> fontDirs, File source) {
        this(wb, doc, fontDirs, FontManager.defaultCacheDir(), source);
    }

    /**
     * 描画の状態を作る。
     *
     * @param wb           変換元のブック
     * @param doc          出力先の PDF
     * @param fontDirs     追加で検索するフォントフォルダ
     * @param fontCacheDir フォントの走査結果のキャッシュを置くフォルダ (null ならキャッシュを使わない)
     * @param source       変換元のファイル
     */
    public RenderContext(XSSFWorkbook wb, PDDocument doc, List<File> fontDirs, File fontCacheDir, File source) {
        this.wb = wb;
        this.doc = doc;
        this.source = source;
        this.fonts = new FontManager(doc, fontDirs, fontCacheDir);
        this.theme = new Theme(wb);
        if (EMPTY_FONT_MSPG) fillEmptyFonts(wb);
        substituteMissingFonts(wb);
        if (CELL_S0 > 0) explicitDefaultStyle(wb);
        syncChartSheets(wb);
        XSSFCellStyle normal = wb.getNumCellStyles() > 0 ? wb.getCellStyleAt(0) : null;
        XSSFFont f = normal != null ? normal.getFont() : wb.getFontAt(0);
        // テーマのないブックは今までどおり cellXfs[0] のフォント (夜間 n0248 など 28 件: 標準スタイルは scheme minor の Calibri、cellXfs[0] は空のフォントで、Excel の標準行は ＭＳ Ｐゴシックの 13.5pt)
        XSSFFont nf = NORMAL_STYLE_FONT && !theme.missing ? normalStyleFont(wb) : null;
        if (nf != null) f = nf;
        this.defaultFontName = f.getFontName();
        // 標準フォントがテーマのフォント (scheme="minor") なら、列幅と標準行の寸法はテーマの日本語フォントで決まる。名前は文字の描画にだけ使う
        // テーマのないブックは、日本語版の Excel の既定のテーマの日本語の書体 (テーマのフォントのセルだけに使う。exp108)
        String minJ = theme.missing ? "游ゴシック" : theme.minorJpan, majJ = theme.missing ? "游ゴシック Light" : theme.majorJpan;
        String metricFont = f.getScheme() == org.apache.poi.ss.usermodel.FontScheme.MINOR
                && minJ != null && !minJ.isEmpty() ? minJ
                : f.getScheme() == org.apache.poi.ss.usermodel.FontScheme.MAJOR
                && majJ != null && !majJ.isEmpty() ? majJ : defaultFontName;
        this.defaultFontSize = f.getFontHeight() / 20.0;
        boolean ja = FontManager.substitute(defaultFontName).startsWith("ipa")
                || (theme.minorEa != null && !theme.minorEa.isEmpty() && FontManager.substitute(theme.minorEa).startsWith("ipa"));
        // 標準のスタイルの名前が「標準」なら日本語版の Excel で作ったブック (組み込みの表示形式などは日本語版のもの。exp96)。
        // 東アジアの既定の書体は、標準の書体が日本語の書体のときだけ標準の書体にする
        this.japanese = ja || normalStyleJa(wb);
        if (ja) fonts.setEastAsianDefault(defaultFontName);
        else if (theme.minorJpan != null) fonts.setEastAsianDefault(theme.minorJpan);
        this.mdw = fonts.maxDigitWidthPx(metricFont, (float) defaultFontSize);
        this.printer = new PrinterModel(fonts, metricFont, defaultFontSize, mdw);
        this.printer.appVersion = appVersion(wb);
        this.printer.nonExcel = nonExcel(wb) && NON_EXCEL_REFIT;
        if (minJ != null && !minJ.isEmpty()) this.printer.minorJpan = minJ;
        if (majJ != null && !majJ.isEmpty()) this.printer.majorJpan = majJ;
        // 英語版の Excel で作ったブック (標準フォント Calibri、テーマの日本語フォント ＭＳ Ｐゴシック) の既定の行高は Calibri の標準行。日本語版はテーマの日本語フォントの標準行にする (exp55)
        if (!metricFont.equals(defaultFontName)) this.printer.latinRow = PrinterModel.measuredScreenRow(defaultFontName, defaultFontSize, 96);
        try {
            var fv = wb.getCTWorkbook().getFileVersion();
            if (fv != null) {
                if (fv.isSetLastEdited()) this.printer.lastEdited = Integer.parseInt(fv.getLastEdited());
                if (fv.isSetLowestEdited()) this.printer.lowestEdited = Integer.parseInt(fv.getLowestEdited());
            }
        } catch (RuntimeException ignored) {
            // 不明
        }
        this.formatter = new ValueFormatter(this.japanese, wb.isDate1904(), wb.getStylesSource());
    }
}
