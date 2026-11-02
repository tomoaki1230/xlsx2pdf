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
    }

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
        double lx = alignWithMargins ? ps.left : 0.5 * 72;
        double rx = pageW - (alignWithMargins ? ps.right : 0.5 * 72);
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
    private record Style(String font, double size, boolean bold, boolean italic, boolean strike, int underline, int shift, Color color) {}

    /** ヘッダーまたはフッターの文字列を描く (ヘッダーは上余白から下へ、フッターは下余白から上へ積む)。 */
    private void paintBlock(PdfCanvas cv, String text, boolean header, double margin, double pageH, double lx, double rx,
                            double fs, int page, int total, String sheetName, String kind, String suffix) throws IOException {
        List<Section> secs = parse(text, fs, page, total, sheetName);
        for (Section s : secs) {
            // 図 (&G) はセクションの先頭に置く
            HeaderImages.Img img = null;
            if (!s.pics.isEmpty()) img = images.get(s.pos + kind + suffix);
            if (img == null && !s.pics.isEmpty()) img = images.get(s.pos + kind);
            List<TextLayout.Line> lines = new ArrayList<>();
            for (List<TextLayout.Run> lr : s.lines) {
                List<TextLayout.Line> l = TextLayout.layout(lr, 0, false, false, ctx.fonts);
                lines.add(l.get(0));
            }
            double total0 = 0;
            for (TextLayout.Line l : lines) {
                if (l.ascent + l.descent == 0) {
                    FontManager.Spec sp = ctx.fonts.spec(ctx.defaultFontName, false, false);
                    l.ascent = (float) (sp.ascent * ctx.defaultFontSize * fs);
                    l.descent = (float) (sp.descent * ctx.defaultFontSize * fs);
                }
                total0 += lineBox(l);
            }
            double imgH = img != null ? img.h() * fs : 0, imgW = img != null ? img.w() * fs : 0;
            double blockH = total0 + imgH;
            // 1 行だけの部分は枠の内側 8 画素 (600 dpi) に置く
            double inset = 0;
            double y = header ? margin + inset : pageH - margin - inset - blockH;
            if (img != null) {
                double ix = s.pos == 'L' ? lx : s.pos == 'R' ? rx - imgW : (lx + rx) / 2 - imgW / 2;
                cv.image(img.image(), ix, y, imgW, imgH, 1);
                y += imgH;
            }
            for (TextLayout.Line l : lines) {
                double side = 0;
                double x = switch (s.pos) {
                    case 'L' -> lx + side;
                    case 'R' -> rx - side - l.width;
                    default -> (lx + rx) / 2 - l.width / 2;
                };
                TextLayout.drawLine(cv, l, x, y + l.ascent + 0, 0, 0);
                y += lineBox(l);
            }
        }
    }

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
        Style st = new Style(fontName, size, bold, italic, strike, underline, shift, color);
        int i = 0;
        while (i <= text.length()) {
            char ch = i < text.length() ? text.charAt(i) : '\0';
            boolean end = i >= text.length();
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
                        fontName = ctx.defaultFontName; size = ctx.defaultFontSize; bold = italic = strike = false;
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
                    case 'D' -> buf.append(LocalDateTime.now().format(DateTimeFormatter.ofPattern(ctx.japanese ? "yyyy/M/d" : "M/d/yyyy")));
                    case 'T' -> buf.append(LocalDateTime.now().format(DateTimeFormatter.ofPattern("H:mm")));
                    case 'F' -> buf.append(ctx.sourceName != null ? ctx.sourceName : "");
                    case 'Z' -> buf.append(ctx.source != null && ctx.source.getAbsoluteFile().getParent() != null
                            ? ctx.source.getAbsoluteFile().getParent() + java.io.File.separator : "");
                    case 'A' -> buf.append(sheetName);
                    case 'G' -> pics.add("G");
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
                        if (!parts[0].equals("-")) fontName = parts[0];
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
                st = new Style(fontName, size, bold, italic, strike, underline, shift, color);
                continue;
            }
            buf.append(ch);
            i++;
        }
        flush(out, pos, lines, pics);
        return out;
    }

    /** 解析したセクションを追加する (空なら追加しない。同じ位置のセクションがあれば続けてつなぐ)。 */
    private static void flush(List<Section> out, char pos, List<List<TextLayout.Run>> lines, List<String> pics) {
        boolean any = !pics.isEmpty();
        for (List<TextLayout.Run> l : lines) if (!l.isEmpty()) any = true;
        if (!any) return;
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
        FontManager.Spec sp = ctx.fonts.spec(st.font(), st.bold(), st.italic());
        TextLayout.Run r = new TextLayout.Run(text, sp, (float) (st.size() * fs), st.color());
        r.fontName = st.font();
        r.ptSize = st.size();
        r.strike = st.strike();
        r.underline = st.underline();
        r.baselineShift = st.shift();
        return r;
    }
}
