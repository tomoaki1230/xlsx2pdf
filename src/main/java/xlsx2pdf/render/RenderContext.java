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

    /**
     * ブックを最後に保存した Excel のメジャーバージョンを返す (docProps/app.xml の AppVersion)。
     *
     * @return 14 = 2010、15 = 2013、16 = 2016 以降、0 = 不明
     */
    static int appVersion(XSSFWorkbook wb) {
        try {
            String v = wb.getProperties().getExtendedProperties().getUnderlyingProperties().getAppVersion();
            if (v != null) return Integer.parseInt(v.substring(0, v.indexOf('.') > 0 ? v.indexOf('.') : v.length()));
        } catch (RuntimeException ignored) {
            // 不明
        }
        return 0;
    }

    /**
     * 描画の状態を作る。
     *
     * @param wb       変換元のブック
     * @param doc      出力先の PDF
     * @param fontDirs 追加で検索するフォントフォルダ
     * @param source   変換元のファイル
     */
    public RenderContext(XSSFWorkbook wb, PDDocument doc, List<File> fontDirs, File source) {
        this.wb = wb;
        this.doc = doc;
        this.source = source;
        this.fonts = new FontManager(doc, fontDirs);
        this.theme = new Theme(wb);
        XSSFCellStyle normal = wb.getNumCellStyles() > 0 ? wb.getCellStyleAt(0) : null;
        XSSFFont f = normal != null ? normal.getFont() : wb.getFontAt(0);
        this.defaultFontName = f.getFontName();
        // 標準フォントがテーマのフォント (scheme="minor") なら、列幅と標準行の寸法はテーマの日本語フォントで決まる。名前は文字の描画にだけ使う
        String metricFont = f.getScheme() == org.apache.poi.ss.usermodel.FontScheme.MINOR
                && theme.minorJpan != null && !theme.minorJpan.isEmpty() ? theme.minorJpan : defaultFontName;
        this.defaultFontSize = f.getFontHeight() / 20.0;
        boolean ja = FontManager.substitute(defaultFontName).startsWith("ipa")
                || (theme.minorEa != null && !theme.minorEa.isEmpty() && FontManager.substitute(theme.minorEa).startsWith("ipa"));
        this.japanese = ja;
        if (ja) fonts.setEastAsianDefault(defaultFontName);
        else if (theme.minorJpan != null) fonts.setEastAsianDefault(theme.minorJpan);
        this.mdw = fonts.maxDigitWidthPx(metricFont, (float) defaultFontSize);
        this.printer = new PrinterModel(fonts, metricFont, defaultFontSize, mdw);
        this.printer.appVersion = appVersion(wb);
        try {
            var fv = wb.getCTWorkbook().getFileVersion();
            if (fv != null) {
                if (fv.isSetLastEdited()) this.printer.lastEdited = Integer.parseInt(fv.getLastEdited());
                if (fv.isSetLowestEdited()) this.printer.lowestEdited = Integer.parseInt(fv.getLowestEdited());
            }
        } catch (RuntimeException ignored) {
            // 不明
        }
        this.formatter = new ValueFormatter(ja, wb.isDate1904(), wb.getStylesSource());
    }
}
