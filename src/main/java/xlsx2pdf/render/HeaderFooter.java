package xlsx2pdf.render;

import java.awt.Color;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHeaderFooter;

/** Excel のヘッダー/フッターの解析と描画。 */
public final class HeaderFooter {
    /** &amp;K の 2 桁のテーマ色番号に対応するテーマの色名。 */
    private static final String[] THEME_COLORS = {"lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"};
    /** 行送り: Excel は文字の箱の下に約 11% の行間を付ける (標準行高と同様)。 */
    static final double LEADING = 1.11;
    /** 和文の書体でないフォントの和文を描く書体。 */
    static final String HF_EA = "ＭＳ Ｐゴシック";

    private final RenderContext ctx;
    private final XSSFSheet sheet;
    private final CTHeaderFooter hf;
    /** ドキュメントに合わせて拡大縮小する / ページ余白に合わせる。 */
    private final boolean scaleWithDoc, alignWithMargins;
    /** &amp;G の図 (セクション ID: LH, CH, RH, LF, CF, RF。先頭ページ・偶数ページは接尾辞付き)。 */
    private final Map<String, HeaderImages.Img> images;

    /** シートのヘッダー/フッターを読み込む。 */
    public HeaderFooter(RenderContext ctx, XSSFSheet sheet) {
        this.ctx = ctx;
        this.sheet = sheet;
        this.hf = sheet.getCTWorksheet().isSetHeaderFooter() ? sheet.getCTWorksheet().getHeaderFooter() : null;
        this.scaleWithDoc = hf == null || !hf.isSetScaleWithDoc() || hf.getScaleWithDoc();
        this.alignWithMargins = hf == null || !hf.isSetAlignWithMargins() || hf.getAlignWithMargins();
        this.images = HeaderImages.load(ctx, sheet);
        double kx = 1, ky = 1;
        try {
            PrinterModel.Factors f = ctx.printer.factors(1.0, 600, 600);
            if (ctx.mdw > 0 && f.mdwP > 0) kx = 0.16 * f.mdwP / ctx.mdw;
            if (f.rowRatio > 0) ky = 0.16 * f.rowRatio;
        } catch (RuntimeException e) {
            // 係数を求められなければ VML の大きさのまま
        }
        this.imgKx = kx;
        this.imgKy = ky;
    }

    /** 画像の大きさに掛ける比 (横・縦)。 */
    private final double imgKx, imgKy;

    /** ヘッダー/フッターの設定がないか。 */
    public boolean isEmpty() { return hf == null; }

    /**
     * 1 ページのヘッダーとフッターを描く。
     *
     * @param pageW            ページの幅 (pt)
     * @param pageH            ページの高さ (pt)
     * @param scale            印刷倍率 (ドキュメントに合わせて拡大縮小するときに使う)
     * @param pageIndexInSheet シート内のページ番号 (0 始まり。先頭ページの判定に使う)
     * @param pageNumber       印刷するページ番号 (&amp;P)
     * @param totalPages       総ページ数 (&amp;N)
     * @param sheetName        シート名 (&amp;A)
     */
    public void paint(PdfCanvas cv, PrintSetup ps, double pageW, double pageH, double scale, int pageIndexInSheet,
                      int pageNumber, int totalPages, String sheetName) throws IOException {
        if (hf == null) return;
        boolean first = pageIndexInSheet == 0 && hf.isSetDifferentFirst() && hf.getDifferentFirst();
        boolean even = !first && hf.isSetDifferentOddEven() && hf.getDifferentOddEven() && pageNumber % 2 == 0;
        String head, foot, suffix;
        if (first) { head = hf.getFirstHeader(); foot = hf.getFirstFooter(); suffix = "FIRST"; }
        else if (even) { head = hf.getEvenHeader(); foot = hf.getEvenFooter(); suffix = "EVEN"; }
        else { head = hf.getOddHeader(); foot = hf.getOddFooter(); suffix = ""; }
        double fs = scaleWithDoc ? scale : 1;
        // 余白に合わせないときは用紙の端から 0.75in (exp121 noalign)
        double lx = alignWithMargins ? ps.left : 0.75 * 72;
        double rx = pageW - (alignWithMargins ? ps.right : 0.75 * 72);
        if (head != null && !head.isEmpty())
            paintBlock(cv, head, true, Math.max(ps.header, ps.minMargin), pageH, lx, rx, fs, pageNumber, totalPages, sheetName, "H", suffix);
        if (foot != null && !foot.isEmpty())
            paintBlock(cv, foot, false, Math.max(ps.footer, ps.minMargin), pageH, lx, rx, fs, pageNumber, totalPages, sheetName, "F", suffix);
    }

    /**
     * 左・中央・右のいずれかのセクション。
     *
     * @param pos   位置 ('L', 'C', 'R')
     * @param lines 行ごとのラン
     * @param pics  図 (&amp;G) の指定
     */
    private record Section(char pos, List<List<TextLayout.Run>> lines, List<String> pics) {}

    /**
     * 文字の書式の状態 (解析中の現在値)。
     *
     * @param underline 0 なし、1 一重、2 二重
     * @param shift     1 上付き、-1 下付き
     */
    private record Style(String font, String ea, double size, boolean bold, boolean italic, boolean strike, int underline, int shift, Color color) {}

    /** ヘッダーまたはフッターの文字列を描く (ヘッダーは上余白から下へ、フッターは下余白から上へ積む)。 */
    private void paintBlock(PdfCanvas cv, String text, boolean header, double margin, double pageH, double lx, double rx,
                            double fs, int page, int total, String sheetName, String kind, String suffix) throws IOException {
        List<Section> secs = parse(text, fs, page, total, sheetName);
        for (Section s : secs) {
            // 図 (&G) はセクションの先頭に置く
            HeaderImages.Img img = null;
            if (!s.pics.isEmpty()) img = images.get(s.pos + kind + suffix);
            if (img == null && !s.pics.isEmpty()) img = images.get(s.pos + kind);
            // 画像の大きさはセルの列幅・行高と同じく「プリンタ / 画面」の比で印刷する (横は数字幅、縦は標準行の比。exp122 img_basic:
            // ＭＳ Ｐゴシック 11 で 150 × 75pt → 138 × 68.04pt)
            double imgH = img != null ? img.h() * fs * imgKy : 0, imgW = img != null ? img.w() * fs * imgKx : 0;
            // 行ごとに: 文字だけ・画像だけ (&G だけの行)・文字と画像 (&G を文字の間に置いた行) に分ける
            List<Row> rows = new ArrayList<>();
            double total0 = 0;
            boolean anyText = false;
            for (List<TextLayout.Run> lr0 : s.lines) {
                int mark = -1;
                for (int k = 0; k < lr0.size(); k++) if (IMG_MARK.equals(lr0.get(k).fontName)) { mark = k; break; }
                List<TextLayout.Run> lr = new ArrayList<>();
                for (TextLayout.Run r : lr0) if (!IMG_MARK.equals(r.fontName)) lr.add(r);
                boolean hasText = false;
                for (TextLayout.Run r : lr) if (!r.text.isEmpty()) hasText = true;
                anyText |= hasText;
                if (mark >= 0 && img != null && !hasText) {
                    rows.add(new Row(null, null, 0, true));
                    continue;
                }
                if (mark >= 0 && img != null) {
                    // 文字と画像: 画像を行の中に置く (exp122 img_text「前の文字 &G 後の文字」)
                    List<TextLayout.Run> left = new ArrayList<>(), right = new ArrayList<>();
                    for (int k = 0; k < lr0.size(); k++) {
                        TextLayout.Run r = lr0.get(k);
                        if (IMG_MARK.equals(r.fontName)) continue;
                        (k < mark ? left : right).add(r);
                    }
                    TextLayout.Line ll = lineOf(left, fs), rl = lineOf(trimTrailingSpaces(right), fs);
                    double pb = HF_PITCH > 0 ? printerPitch(lr, fs) : 0;
                    if (pb <= 0) pb = lineBox(ll.ascent >= rl.ascent ? ll : rl);
                    Row row = new Row(ll, rl, pb, true);
                    rows.add(row);
                    continue;
                }
                // 収まる行は折り返さずに置く。行末の全角の空白は幅に数え (n0996 の「事業所名：」+ 全角の空白)、半角の空白は落とす (exp121 spaces)
                lr = trimTrailingSpaces(lr);
                List<TextLayout.Line> ls = TextLayout.layout(lr, 0, false, false, ctx.fonts);
                if (HF_WRAP && !ls.isEmpty() && ls.get(0).width > rx - lx + 0.01) {
                    List<TextLayout.Line> w = TextLayout.layout(lr, (float) (rx - lx), true, false, ctx.fonts);
                    if (!w.isEmpty()) ls = w;
                }
                double pb = HF_PITCH > 0 ? printerPitch(lr, fs) : 0;
                for (TextLayout.Line l : ls) {
                    fillEmpty(l, fs);
                    rows.add(new Row(l, null, pb > 0 ? pb : lineBox(l), false));
                    if (!HF_WRAP) break;
                }
            }
            if (img != null && !anyText) rows.removeIf(r -> !r.image());   // 画像だけの部分は文字の行を数えない (exp122 img_basic)
            if (img != null && rows.stream().noneMatch(Row::image)) rows.add(0, new Row(null, null, 0, true));   // &G のない画像は先頭に
            // 行の高さ: 画像だけの行は画像の高さ (後ろに文字の行が続くときは 1.17pt 空ける)、文字と画像の行は画像の高さ + 文字の行の下の部分
            List<Double> adv = new ArrayList<>();
            for (int k = 0; k < rows.size(); k++) {
                Row r = rows.get(k);
                double h;
                if (r.image() && r.line() == null) h = imgH + (k + 1 < rows.size() ? IMG_GAP * fs : 0);
                else if (r.image()) h = Math.max(r.pitch(), imgH + r.pitch() - Math.max(r.line().ascent, r.right().ascent));
                else h = r.pitch();
                adv.add(h);
                total0 += h;
            }
            double blockH = total0;
            // 1 行だけの部分は枠の内側 8 画素 (600 dpi) に置く
            // 書式が途中で変わる行 (ランが 2 つ以上) は内側にずらさない (実ファイルの 77 行、exp121 supsub、exp102 codes の右)
            boolean textOnly = img == null && rows.size() == 1 && !rows.get(0).image();
            boolean oneRun = textOnly && runCount(rows.get(0).line()) <= 1;
            boolean single = textOnly && (HF_ONE_RUN < 2 || oneRun);
            double inset = (HF_INSET & 1) != 0 && single ? HF_INSET_PT : 0;
            boolean sideOk = textOnly && (HF_ONE_RUN < 1 || oneRun);
            double side = sideOk && (s.pos == 'L' && (HF_INSET & 2) != 0 || s.pos == 'R' && (HF_INSET & 4) != 0) ? HF_INSET_PT : 0;
            double y = header ? margin + inset : pageH - margin - inset - blockH;
            for (int k = 0; k < rows.size(); k++) {
                Row r = rows.get(k);
                if (r.image() && r.line() == null) {
                    double ix = s.pos == 'L' ? lx : s.pos == 'R' ? rx - imgW : (lx + rx) / 2 - imgW / 2;
                    cv.image(img.image(), ix, y, imgW, imgH, 1);
                } else if (r.image()) {
                    double w = r.line().width + imgW + r.right().width;
                    double x = s.pos == 'L' ? lx : s.pos == 'R' ? rx - w : (lx + rx) / 2 - w / 2;
                    double asc = Math.max(r.line().ascent, r.right().ascent);
                    double base = y + Math.max(asc, imgH);
                    TextLayout.drawLine(cv, r.line(), x, base, 0, 0);
                    cv.image(img.image(), x + r.line().width, base - imgH + IMG_INLINE_DROP * fs, imgW, imgH, 1);
                    TextLayout.drawLine(cv, r.right(), x + r.line().width + imgW, base, 0, 0);
                } else {
                    TextLayout.Line l = r.line();
                    double x = switch (s.pos) {
                        case 'L' -> lx + side;
                        case 'R' -> rx - side - l.width;
                        default -> (lx + rx) / 2 - l.width / 2;
                    };
                    TextLayout.drawLine(cv, l, x, y + l.ascent + supExtra(l) + HF_BASE + (explicitLook(l) ? HF_CODE_DROP : 0), 0, 0);
                }
                y += adv.get(k);
            }
        }
    }

    /**
     * ヘッダー/フッターの 1 行。
     *
     * @param line  文字の行 (文字と画像の行では画像の前の文字。画像だけの行は null)
     * @param right 文字と画像の行の、画像の後の文字
     * @param image 画像を含む行
     */
    private record Row(TextLayout.Line line, TextLayout.Line right, double pitch, boolean image) {}

    /** 画像の位置を表す印のラン (fontName)。 */
    static final String IMG_MARK = "\u0001&G";
    /** 画像だけの行の後に文字の行が続くときの間 (pt。exp122 img_text のフッター)。 */
    static final double IMG_GAP = 1.17;
    /** 文字と画像の行で、画像の下端を基線より下げる量 (pt。exp122 img_text のヘッダー)。 */
    static final double IMG_INLINE_DROP = 1.32;

    /** 折り返さない 1 行 (空なら既定の書体の高さ)。 */
    private TextLayout.Line lineOf(List<TextLayout.Run> runs, double fs) {
        List<TextLayout.Line> ls = TextLayout.layout(runs, 0, false, false, ctx.fonts);
        if (ls.isEmpty()) ls = TextLayout.layout(List.of(mkRun("", new Style(ctx.defaultFontName, null, ctx.defaultFontSize, false, false, false, 0, 0, Color.BLACK), fs)), 0, false, false, ctx.fonts);
        TextLayout.Line l = ls.get(0);
        fillEmpty(l, fs);
        return l;
    }

    /** 文字のない行の高さを既定の書体で決める。 */
    private void fillEmpty(TextLayout.Line l, double fs) {
        if (l.ascent + l.descent == 0) {
            FontManager.Spec sp = ctx.fonts.spec(ctx.defaultFontName, false, false);
            l.ascent = (float) (sp.ascent * ctx.defaultFontSize * fs);
            l.descent = (float) (sp.descent * ctx.defaultFontSize * fs);
        }
    }

    /** 行送りをプリンタの標準行にする (exp15)。 */
    static final int HF_PITCH = 1;
    /** 内側の余白: 1 = 1 行の部分を縦に、2 = 左の部分を横に、4 = 右の部分を横に (右は長い文字で合わないので使わない)。 */
    static final int HF_INSET = 7;
    static final double HF_INSET_PT = 0.96;
    static final boolean HF_TRIM = true;
    static final boolean HF_SEPARATE = true;
    static final boolean HF_WRAP = true;
    static final double HF_BASE = 0;

    /**
     * 行送り (pt): Excel はその行の文字の書体・大きさのプリンタの標準行 (600 dpi、printrows.tsv) の最大で行を送る (exp15)。
     * ドキュメントに合わせて拡大縮小するときは倍率を掛ける。求められなければ 0。
     */
    private double printerPitch(List<TextLayout.Run> runs, double fs) {
        int best = 0;
        List<TextLayout.Run> rs = runs;
        if (rs.isEmpty()) rs = List.of(mkRun(" ", new Style(ctx.defaultFontName, null, ctx.defaultFontSize, false, false, false, 0, 0, Color.BLACK), fs));
        for (TextLayout.Run r : rs) {
            if (r.fontName == null || r.ptSize <= 0) return 0;
            String n = FontManager.norm(r.fontName);
            // ＭＳ 系はどの大きさでも p + 10 (printrows.tsv の 6〜28pt がすべてこの形。72pt も Excel のフッターの位置と合う。exp121 sizes)
            int px = n.startsWith("ms ") ? (int) Math.round(r.ptSize * 600 / 72.0) + 10 : ctx.printer.autoRowDev(r.fontName, r.ptSize, 1.0, 600);
            if (px <= 0) return 0;
            best = Math.max(best, px);
        }
        return best * 72.0 / 600 * fs;
    }

    /**
     * 上付きの文字の上端が行の文字の上端より上に出るとき、Excel は行をその分 (から 0.43pt 引いた量) 下げる (exp121 supsub:
     * 11pt で 1.59pt、14pt で 2.25pt)。
     */
    private static double supExtra(TextLayout.Line l) {
        if (!HF_CELLSHIFT) return 0;
        double top = 0;
        for (TextLayout.Run r : l.runs) {
            if (r.baselineShift <= 0 || r.text.isEmpty()) continue;
            top = Math.max(top, r.shift() + r.spec.ascent * r.effSize());
        }
        return Math.max(0, top - l.ascent - SUP_EXTRA_PAD);
    }

    static final boolean HF_CELLSHIFT = true;
    static final double SUP_EXTRA_PAD = 0.43;

    /** 行末の半角の空白を落としたランの並び (元のランは変えない)。 */
    private static List<TextLayout.Run> trimTrailingSpaces(List<TextLayout.Run> lr) {
        int k = lr.size() - 1;
        while (k >= 0 && lr.get(k).text.isEmpty()) k--;
        if (k < 0 || !lr.get(k).text.endsWith(" ")) return lr;
        List<TextLayout.Run> out = new ArrayList<>(lr);
        while (k >= 0) {
            TextLayout.Run r = out.get(k);
            String t = r.text.replaceAll(" +$", "");
            if (t.length() == r.text.length()) break;
            TextLayout.Run c = r.copy(t);
            out.set(k, c);
            if (!t.isEmpty()) break;
            k--;
        }
        return out;
    }

    /** 書体か大きさを指定した (&amp;"書体"・&amp;数字) 文字のある行か。 */
    private boolean explicitLook(TextLayout.Line l) {
        for (TextLayout.Run r : l.runs) {
            if (r.text.isEmpty() || IMG_MARK.equals(r.fontName)) continue;
            if (r.ptSize != ctx.defaultFontSize || r.fontName != null && !r.fontName.equals(ctx.defaultFontName)) return true;
        }
        return false;
    }

    /** 書体か大きさを指定した行の基線を下げる量 (pt)。実ファイル 376 件の帯の比較で、指定のある行だけ本ツールが約 0.2〜0.3pt 高い (exp15 は一様に約 0.3)。 */
    static final double HF_CODE_DROP = 0.25;

    /** 書式の違うランの数 (&amp;P・&amp;D などの項目で区切られただけの同じ書式のランは 1 つと数える。n0509「出力日：&amp;D」)。 */
    private static int runCount(TextLayout.Line l) {
        int n = 0;
        TextLayout.Run prev = null;
        for (TextLayout.Run r : l.runs) {
            if (r.text.isEmpty()) continue;
            if (prev == null || !sameLook(prev, r)) n++;
            prev = r;
        }
        return n;
    }

    private static boolean sameLook(TextLayout.Run a, TextLayout.Run b) {
        return a.spec == b.spec && a.size == b.size && java.util.Objects.equals(a.color, b.color) && a.underline == b.underline
                && a.strike == b.strike && a.baselineShift == b.baselineShift;
    }

    /** 内側の余白 (縦・横) は、書式が途中で変わらない 1 行の部分だけ (実ファイル 376 件の帯の比較で位置の合う行 2735 → 3424、exp102 codes)。 */
    static final int HF_ONE_RUN = 2;

    /** 行送り (pt)。 */
    static double lineBox(TextLayout.Line l) {
        return (l.ascent + l.descent) * LEADING;
    }

    /**
     * ヘッダー/フッターの書式コード (&amp;L &amp;C &amp;R、&amp;P &amp;N、&amp;"フォント,スタイル"、&amp;数字 など) を解析する。
     *
     * @param fs 文字サイズの倍率
     * @return セクションの一覧
     */
    private List<Section> parse(String text, double fs, int page, int total, String sheetName) {
        List<Section> out = new ArrayList<>();
        String fontName = ctx.defaultFontName;
        String eaName = null;   // 和文の書体 (null なら HF_EA。書体名「-」はテーマの和文の書体)
        double size = ctx.defaultFontSize;
        boolean bold = false, italic = false, strike = false;
        int underline = 0, shift = 0;
        Color color = Color.BLACK;
        char pos = 'C';
        List<List<TextLayout.Run>> lines = new ArrayList<>();
        List<String> pics = new ArrayList<>();
        lines.add(new ArrayList<>());
        StringBuilder buf = new StringBuilder();
        // バッファの文字列に適用する書式
        Style st = new Style(fontName, eaName, size, bold, italic, strike, underline, shift, color);
        int i = 0;
        while (i <= text.length()) {
            char ch = i < text.length() ? text.charAt(i) : '\0';
            boolean end = i >= text.length();
            // 末尾の単独の & は Excel は描かない (exp121 amp)
            if (!end && ch == '&' && i + 1 >= text.length()) { i++; continue; }
            boolean code = !end && ch == '&' && i + 1 < text.length();
            boolean newline = !end && ch == '\n';
            if (end || code || newline) {
                if (buf.length() > 0) {
                    lines.get(lines.size() - 1).add(mkRun(buf.toString(), st, fs));
                    buf.setLength(0);
                }
                if (end) break;
                if (newline) { lines.add(new ArrayList<>()); i++; continue; }
                char c = text.charAt(i + 1);
                i += 2;
                switch (c) {
                    case 'L', 'C', 'R' -> {
                        flush(out, pos, lines, pics);
                        lines = new ArrayList<>();
                        lines.add(new ArrayList<>());
                        pics = new ArrayList<>();
                        pos = c;
                        // Excel はセクションごとに書式を戻す
                        fontName = ctx.defaultFontName; eaName = null; size = ctx.defaultFontSize; bold = italic = strike = false;
                        underline = 0; shift = 0; color = Color.BLACK;
                    }
                    case '&' -> buf.append('&');
                    case 'P', 'N' -> {
                        int base = c == 'P' ? page : total;
                        int k = i;
                        if (k < text.length() && (text.charAt(k) == '+' || text.charAt(k) == '-')) {
                            int j = k + 1;
                            while (j < text.length() && Character.isDigit(text.charAt(j))) j++;
                            if (j > k + 1) {
                                int d = Integer.parseInt(text.substring(k + 1, j));
                                base += text.charAt(k) == '+' ? d : -d;
                                i = j;
                            }
                        }
                        buf.append(base);
                    }
                    case 'D' -> buf.append(LocalDateTime.now().format(DateTimeFormatter.ofPattern(ctx.japanese ? "yyyy/MM/dd" : "M/d/yyyy")));
                    case 'T' -> buf.append(LocalDateTime.now().format(DateTimeFormatter.ofPattern("H:mm")));
                    case 'F' -> buf.append(ctx.sourceName != null ? ctx.sourceName : "");
                    case 'Z' -> buf.append(ctx.source != null && ctx.source.getAbsoluteFile().getParent() != null
                            ? ctx.source.getAbsoluteFile().getParent() + java.io.File.separator : "");
                    case 'A' -> buf.append(sheetName);
                    case 'G' -> {
                        pics.add("G");
                        TextLayout.Run m = mkRun("\uFFFC", st, fs);
                        m.fontName = IMG_MARK;
                        lines.get(lines.size() - 1).add(m);
                    }
                    case 'B' -> bold = !bold;
                    case 'I' -> italic = !italic;
                    case 'U' -> underline = underline == 1 ? 0 : 1;
                    case 'E' -> underline = underline == 2 ? 0 : 2;
                    case 'S' -> strike = !strike;
                    case 'X' -> shift = shift == 1 ? 0 : 1;
                    case 'Y' -> shift = shift == -1 ? 0 : -1;
                    case 'K' -> {
                        if (i + 6 <= text.length()) {
                            String hex = text.substring(i, i + 6);
                            if (hex.matches("[0-9A-Fa-f]{6}")) {
                                color = Theme.hex(hex);
                                i += 6;
                            } else if (hex.matches("\\d\\d[+-]\\d{3}")) {
                                int t = Integer.parseInt(hex.substring(0, 2));
                                Color base = t < THEME_COLORS.length ? ctx.theme.scheme.get(THEME_COLORS[t]) : Color.BLACK;
                                double tint = Integer.parseInt(hex.substring(3)) / 100.0 * (hex.charAt(2) == '-' ? -1 : 1);
                                color = Theme.applyTint(base == null ? Color.BLACK : base, tint);
                                i += 6;
                            }
                        }
                    }
                    case '"' -> {
                        int j = text.indexOf('"', i);
                        if (j < 0) j = text.length();
                        String spec = text.substring(i, j);
                        i = Math.min(text.length(), j + 1);
                        String[] parts = spec.split(",", 2);
                        // 「-」は書体名を指定しない形で、Excel はテーマの本文の書体で描く (exp102: 標準が Arial・メイリオ・ＭＳ Ｐ明朝でも游ゴシック)
                        if (!parts[0].equals("-")) { fontName = parts[0]; eaName = null; }
                        else {
                            if (ctx.theme.minorLatin != null && !ctx.theme.minorLatin.isEmpty()) fontName = ctx.theme.minorLatin;
                            // 和文はテーマの和文の書体 (exp121 fontstyle: Aptos Narrow のテーマで和文は游ゴシック)
                            String ea = !ctx.theme.minorEa.isEmpty() ? ctx.theme.minorEa : ctx.theme.minorJpan;
                            eaName = ea != null && !ea.isEmpty() ? ea : null;
                        }
                        String style = parts.length > 1 ? parts[1].toLowerCase() : "";
                        if (!style.isEmpty()) {
                            bold = style.contains("bold") || style.contains("太字");
                            italic = style.contains("italic") || style.contains("oblique") || style.contains("斜体");
                        }
                    }
                    default -> {
                        if (Character.isDigit(c)) {
                            int j = i;
                            while (j < text.length() && Character.isDigit(text.charAt(j))) j++;
                            size = Integer.parseInt(text.substring(i - 1, j));
                            i = j;
                        }
                    }
                }
                st = new Style(fontName, eaName, size, bold, italic, strike, underline, shift, color);
                continue;
            }
            if (ch != '\t' && ch != '\r') buf.append(ch);   // タブは幅 0 (exp122 tabs)。CR は改行 (LF) の前のものなので捨てる
            i++;
        }
        flush(out, pos, lines, pics);
        return out;
    }

    /**
     * 解析したセクションを追加する (空なら追加しない)。末尾の空行は数えない (exp121 blanklines)。
     * 同じ位置のセクションが繰り返されたときは、Excel はつながずに別々に同じ位置へ重ねて描く (exp121 order)。
     */
    private static void flush(List<Section> out, char pos, List<List<TextLayout.Run>> lines, List<String> pics) {
        boolean any = !pics.isEmpty();
        for (List<TextLayout.Run> l : lines) if (!l.isEmpty()) any = true;
        if (!any) return;
        while (HF_TRIM && lines.size() > 1 && lines.get(lines.size() - 1).isEmpty()) lines.remove(lines.size() - 1);
        if (HF_SEPARATE) { out.add(new Section(pos, lines, pics)); return; }
        // 同じ位置の既存セクションにつなぐ
        for (Section s : out) {
            if (s.pos == pos) {
                s.lines.get(s.lines.size() - 1).addAll(lines.get(0));
                for (int k = 1; k < lines.size(); k++) s.lines.add(lines.get(k));
                s.pics.addAll(pics);
                return;
            }
        }
        out.add(new Section(pos, lines, pics));
    }

    /** 書式を適用したランを作る。 */
    private TextLayout.Run mkRun(String text, Style st, double fs) {
        // 和文の書体でないフォントの和文は、セル (ＭＳ ゴシック) と違い ＭＳ Ｐゴシックで描く (exp102: Arial のヘッダー・フッター。テーマの書体によらない)
        FontManager.Spec sp = FontManager.substitute(st.font()).startsWith("ipa") ? ctx.fonts.spec(st.font(), st.bold(), st.italic())
                : ctx.fonts.spec(st.font(), st.ea() != null ? st.ea() : HF_EA, st.bold(), st.italic());
        TextLayout.Run r = new TextLayout.Run(text, sp, (float) (st.size() * fs), st.color());
        // 行送りを決める書体: テーマの和文の書体で描く和文があれば、その書体 (exp121 fontstyle の「-,Bold」)
        r.fontName = st.ea() != null && text.chars().anyMatch(ch -> ch >= 0x3000) ? st.ea() : st.font();
        r.ptSize = st.size();
        r.strike = st.strike();
        r.underline = st.underline();
        r.baselineShift = st.shift();
        // 上付き・下付きの大きさとずらし量はセルと同じ (66% を整数に丸め、上 0.497 倍・下 0.198 倍。exp121 supsub)
        r.cellShift = HF_CELLSHIFT;
        return r;
    }
}
