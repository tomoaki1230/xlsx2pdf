package xlsx2pdf.render;

/**
 * Excel が印刷するときの列幅・行高の換算を再現する。
 *
 * <p>Excel は印刷ページをプリンタのピクセル単位でレイアウトする。標準フォントを、拡大縮小後のサイズ
 * (フォントサイズ × 印刷倍率) とプリンタ解像度 (pageSetup の horizontalDpi/verticalDpi。既定 600 dpi) で
 * 実体化し、列幅はその実体化したフォントの最大数字幅を使う通常の「文字数 → ピクセル」の式で、行高は
 * 実体化した標準行高と画面の標準行高の比で換算する。Excel が出力した PDF で実測して合わせている。
 *
 * <p>主な記号: p = 実体化したフォントの ppem (デバイスピクセル)、mdw = 最大数字幅。
 */
public final class PrinterModel {
    /** 標準フォントの最大数字送り幅 (em)。 */
    final double digitEm;
    /** 標準フォントの win アセント・ディセント (em)。 */
    final double asc, desc;
    /** 標準フォントのサイズ (pt)。 */
    final double size;
    /** 画面 (96 dpi) の最大数字幅 (px)。 */
    final int screenMdw;
    /** 画面の標準行高 (px)。 */
    final int screenRow;
    /** ＭＳ ゴシック/明朝系 (em の高さ = アセント + ディセント)。 */
    final boolean ms;
    /** 游ゴシック (実体化した高さを round(1.577 p - 4.77) で近似)。 */
    final boolean yu;

    /** 作成者の環境に合わせる (--author-env)。false なら今の Excel・表示倍率 100% の PC に合わせる。 */
    public boolean authorEnv;
    /** 標準フォントの名前 (テーマの日本語フォントと別の書体) の画面の標準行 (px)。同じ書体なら 0。 */
    public int latinRow;
    /** Excel 以外のアプリで保存したブック: 手で決めていない行の高さ (ht) は Excel が開くときに測り直す。 */
    public boolean nonExcel;
    /** テーマのフォント (scheme minor / major) の日本語の書体。行の高さはこの書体で決まる (日本語版の Excel)。 */
    public String minorJpan, majorJpan;
    /** ブックを最後に保存した Excel のバージョン (0 = 不明)。 */
    public int appVersion;
    /** ブックの fileVersion の lastEdited / lowestEdited (Excel の世代。7 = 2016 以降、0 = 不明)。 */
    public int lastEdited, lowestEdited;

    private final FontManager fm;
    private final String fontName;
    /** {@link #autoRowDev} の結果のキャッシュ。 */
    private final java.util.Map<String, Integer> autoCache = new java.util.HashMap<>();

    /**
     * 標準フォントの印刷モデルを作る。
     *
     * @param fontName  標準フォント名
     * @param size      標準フォントのサイズ (pt)
     * @param screenMdw 画面 (96 dpi) の最大数字幅 (px)
     */
    /** フォントの管理 (折り返しの行数を数えるため)。 */
    FontManager fontManager() { return fm; }

    public PrinterModel(FontManager fm, String fontName, double size, int screenMdw) {
        this.fm = fm;
        this.fontName = fontName;
        this.size = size;
        this.screenMdw = screenMdw;
        String n = FontManager.norm(fontName);
        FontManager.Spec sp = fm.spec(fontName, false, false);
        ms = isMs(n);
        yu = n.startsWith("游ゴシック") || n.startsWith("yu gothic");
        double[] g = FontManager.gdiMetrics(fontName);
        if (g != null) digitEm = g[2];  // 実フォントの最大数字送り幅 (ＭＳ ゴシック 0.5、游ゴシック 1139/2048 など)
        else if (ms) digitEm = 0.5;
        else {
            double max = 0;
            for (char c = '0'; c <= '9'; c++) max = Math.max(max, sp.face.advance(c));
            digitEm = max;
        }
        Double re = ROW_EM.get(n);
        rowEm = re != null ? re : 0;
        asc = sp.ascent;
        desc = sp.descent;
        int known = knownScreenRow(n, size);
        long p96 = Math.round(size * 96 / 72);
        // 游ゴシックは round(ppem × 5/3) (10pt → 22 px、11pt → 25 px)、ＭＳ 系は p + p/11 + 2 (12pt → 19 px = ファイルに見られる 14.25pt)
        screenRow = known > 0 ? known
                : yu ? (int) Math.round(p96 * 5 / 3.0)
                : ms ? msRow(p96)
                : screenRowAt(p96);
    }

    /**
     * プリンタの標準行のセルの高さ (em) の実測値: h = round(k × ppem) として行の式に使う。real-convert の測定用ブック
     * (倍率 85/70/55% のシート、96 dpi) で、高さを指定した行の印刷から逆算して当てはめた。
     * フォントの winAscent + winDescent とは一致しない (Calibri 1.221 に対し 1.174)。
     */
    static final java.util.Map<String, Double> ROW_EM = loadRowEm();

    /** プリンタの標準行のセルの高さの表 (src/main/resources/rowem.tsv) を読む。 */
    private static java.util.Map<String, Double> loadRowEm() {
        java.util.Map<String, Double> m = new java.util.HashMap<>();
        try (var in = PrinterModel.class.getResourceAsStream("/rowem.tsv")) {
            if (in == null) return m;
            for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\\n")) {
                if (l.startsWith("#") || l.isBlank()) continue;
                String[] f = l.split("\\t");
                m.put(f[0], Double.parseDouble(f[1].trim()));
            }
        } catch (Exception e) {
            // 表がなければフォントのアセント + ディセントで求める
        }
        return m;
    }
    /** プリンタの標準行のセルの高さ (em。表になければ 0)。 */
    private double rowEm;

    /** プリンタの標準行の実測表 (正規化したフォント名|大きさ → 600 dpi・倍率 100% のデバイスピクセル)。 */
    private static final java.util.Map<String, Integer> PROW = loadPrintRows();

    private static java.util.Map<String, Integer> loadPrintRows() {
        java.util.Map<String, Integer> m = new java.util.HashMap<>();
        try (var in = PrinterModel.class.getResourceAsStream("/printrows.tsv")) {
            if (in == null) return m;
            for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\n")) {
                if (l.startsWith("#") || l.isBlank()) continue;
                String[] f = l.split("\t");
                m.put(f[0] + "|" + Double.parseDouble(f[1]), Integer.parseInt(f[2].trim()));
            }
        } catch (Exception e) {
            // 表がなければ式で求める
        }
        return m;
    }

    /** 倍率 1・600 dpi のプリンタの標準行の実測値 (表になければ 0)。 */
    static int measuredPrintRow(String fontName, double size, double scale, int dpi) {
        if (scale != 1 || dpi != 600 || fontName == null) return 0;
        String key = FontManager.norm(fontName);
        Integer v = PROW.get(key + "|" + size);
        if (v != null) return v;
        if (isMs(key)) return 0; // MS 系は式 (msRow) が合う
        // 表にない大きさは、同じフォントの最も近い大きさの値から比例で求める (Times New Roman 16pt: 145 × 16 / 14 = 165)
        double best = -1;
        int bv = 0;
        for (var e : PROW.entrySet()) {
            int bar = e.getKey().lastIndexOf('|');
            if (!e.getKey().substring(0, bar).equals(key)) continue;
            double sz = Double.parseDouble(e.getKey().substring(bar + 1));
            if (best < 0 || Math.abs(sz - size) < Math.abs(best - size)) { best = sz; bv = e.getValue(); }
        }
        return best > 0 ? (int) Math.floor(bv * size / best) : 0;
    }

    /** 画面の標準行の実測表 (正規化したフォント名|大きさ|DPI → px)。 */
    private static final java.util.Map<String, Integer> SROW = loadScreenRows();
    /** 画面の標準行の表にある書体名 (正規化した名前。Windows の Excel が使える書体: クラウドフォントを含む)。 */
    private static final java.util.Set<String> SROW_FONTS = new java.util.HashSet<>();
    static {
        for (String k : SROW.keySet()) SROW_FONTS.add(k.substring(0, k.indexOf('|')));
    }

    private static java.util.Map<String, Integer> loadScreenRows() {
        java.util.Map<String, Integer> m = new java.util.HashMap<>();
        try (var in = PrinterModel.class.getResourceAsStream("/screenrows.tsv")) {
            if (in == null) return m;
            for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\n")) {
                if (l.startsWith("#") || l.isBlank()) continue;
                String[] f = l.split("\t");
                m.put(f[0] + "|" + Double.parseDouble(f[1]) + "|" + Integer.parseInt(f[2]), Integer.parseInt(f[3].trim()));
            }
        } catch (Exception e) {
            // 表がなければ式で求める
        }
        return m;
    }

    /**
     * Excel が画面 DPI dpi で使う標準行 (px) の実測値 (real-convert の測定用ブックの standardHeight。ＭＳ 系・游ゴシック・メイリオ・
     * BIZ UD・Calibri など 18 書体 × 9〜14pt × 96/120/144/168 dpi)。表になければ 0。
     */
    static int measuredScreenRow(String fontName, double size, int dpi) {
        if (fontName == null) return 0;
        Integer v = SROW.get(FontManager.norm(fontName) + "|" + size + "|" + dpi);
        if (v != null) return v;
        return FontCalc.screenRow(fontName, size, dpi);
    }

    /** 画面の最大数字幅の実測表 (正規化したフォント名|大きさ|DPI → px)。 */
    private static final java.util.Map<String, Integer> SMDW = loadTable("/screenmdw.tsv");

    private static java.util.Map<String, Integer> loadTable(String res) {
        java.util.Map<String, Integer> m = new java.util.HashMap<>();
        try (var in = PrinterModel.class.getResourceAsStream(res)) {
            if (in == null) return m;
            for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\n")) {
                if (l.startsWith("#") || l.isBlank()) continue;
                String[] f = l.split("\t");
                m.put(f[0] + "|" + Double.parseDouble(f[1]) + "|" + Integer.parseInt(f[2]), Integer.parseInt(f[3].trim()));
            }
        } catch (Exception e) {
            // 表がなければ計算で求める
        }
        return m;
    }

    /**
     * Excel が画面 DPI dpi で使う最大数字幅 (px、列幅の基準) の実測値 (real-convert の測定用ブックの列幅から逆算。
     * 56 書体 × 9〜14pt × 96/120/144/168 dpi)。表になければ 0。フォントの数字の送り幅 × ppem の四捨五入とは
     * ヒンティングで食い違うことがある (Times New Roman 11pt・96 dpi: 計算 8、Excel 7)。
     */
    public static int measuredMdw(String fontName, double size, int dpi) {
        if (fontName == null) return 0;
        Integer v = SMDW.get(FontManager.norm(fontName) + "|" + size + "|" + dpi);
        if (v != null) return v;
        return FontCalc.digitWidth(fontName, size, dpi);
    }

    /** よく使うフォントで Excel が使う画面の標準行高 (px)。該当しなければ 0。 */
    private int knownScreenRow(String n, double size) {
        int m = measuredScreenRow(fontName, size, 96);
        if (m > 0) return m;
        if (Math.abs(size - 11) < 0.01) {
            if (n.startsWith("calibri")) return 20;
            if (n.startsWith("游") || n.startsWith("yu ")) return 25;
            if (n.startsWith("meiryo ui")) return 21;
            if (n.startsWith("メイリオ") || n.startsWith("meiryo")) return 25;
        }
        return 0;
    }

    /** 実体化したフォントの高さ (px)。ＭＳ 系は ppem、游ゴシックは近似式、その他はアセント + ディセント。 */
    private long fontHeight(long p) {
        return ms ? p : yu ? yuH(p) : Math.round(asc * p) + Math.round(desc * p);
    }

    /** 画面 (96 dpi) の標準行 (px)。 */
    public int screenRowPx() { return screenRow; }

    /** ppem p のプリンタの標準行 (デバイスピクセル)。 */
    public int printerRowPx(long p) { return rowPx(p); }

    /**
     * ppem p で実体化したフォントのプリンタの標準行高 (デバイスピクセル。プリンタには枠線のピクセルがない)。
     * 実測: ＭＳ Ｐゴシック 11pt は 600 dpi/100% で 102、86% で 88、82% で 83、300 dpi/92% で 47。
     */
    int rowPx(long p) {
        if (ms) return msRow(p);
        long h = rowEm > 0 ? Math.round(rowEm * p) : fontHeight(p);
        if (yu) return (int) (h + h / 11 + 2);  // 游ゴシックも ＭＳ 系と同じ行間の規則
        return (int) (h + (h + 4) / 10 + 1); // p/10 の四捨五入 (0.5 は切り捨て): 75→7, 79→8, 92→9, 42→4
    }

    /**
     * ＭＳ ゴシック/明朝系を ppem p で実体化したときのプリンタの標準行高: p + floor(p/11) + 2 (行間は 11 ppem ごとに
     * 1 px 増える)。Excel の PDF で p = 18〜96 を実測 (45→51, 66→74, 86→95, 92→102, 96→106)。
     */
    static int msRow(long p) {
        return (int) (p + p / 11 + 2);
    }

    /**
     * 標準フォントを ppem p で実体化したときの画面の標準行高 (px)。画面は従来の GDI の行 (h + (h+4)/10 + 1) に
     * 枠線の 1 px を足したもの: 96 dpi 18 px、120 dpi 22 px、144 dpi 26 px、192 dpi 34 px
     * (= ファイルに見られる既定の高さ 13.5 / 13.2 / 13.0 / 12.75pt)。
     */
    int screenRowAt(long p) {
        long h = fontHeight(p);
        return (int) (h + (h + 4) / 10 + 1) + 1;
    }

    /** 游ゴシックを ppem p で実体化したときの高さ (Excel の PDF で近似)。 */
    static long yuH(long p) {
        return Math.round(1.577 * p - 4.77);
    }

    /** 正規化したフォント名 ({@link FontManager#norm}) が ＭＳ ゴシック/明朝系か。 */
    static boolean isMs(String norm) {
        // ＭＳ ゴシック・明朝の仲間だけ (MS Sans Serif・MS Shell Dlg などは含めない)
        return norm.startsWith("ms") && (norm.contains("ゴシック") || norm.contains("明朝")
                || norm.contains("gothic") || norm.contains("mincho"));
    }

    /**
     * 最大のフォントが指定のものである自動調整行の、印刷時の高さ (デバイスピクセル) を返す。Excel は自動調整行を、
     * 保存された高さを拡大縮小するのではなく、実体化したプリンタのフォントから計算し直す。
     *
     * @param font   フォント名 (null なら標準フォント)
     * @param sizePt フォントサイズ (pt)
     * @param scale  印刷倍率 (1.0 = 100%)
     * @param vdpi   縦の解像度
     */
    /**
     * 折り返した文字の行送り (プリンタの画素)。字面の高さが em と同じフォントは 100% で ppem + 10 画素、縮小時はそれを倍率で丸める
     * (Excel の行送り: 8pt 77、10pt 93、11pt 102、12pt 110、14pt 127、16pt 143。10pt × 77% は 72)。ほかのフォントは {@link #autoRowDev}。
     */
    public int linePitchDev(String font, double sizePt, double scale, int vdpi) {
        String n = FontManager.norm(font == null ? fontName : font);
        if (!authorEnv && vdpi == 600 && (isMs(n) || FontCalc.emHeight(font == null ? fontName : font))) {
            long p = Math.round(sizePt * vdpi / 72.0);
            return (int) Math.floor((p + 10) * scale + 0.5 + 1e-9);
        }
        return autoRowDev(font, sizePt, scale, vdpi);
    }

    public int autoRowDev(String font, double sizePt, double scale, int vdpi) {
        String key = font + "|" + sizePt + "|" + scale + "|" + vdpi;
        Integer v = autoCache.get(key);
        if (v != null) return v;
        int mr = measuredPrintRow(font == null ? fontName : font, sizePt, scale, vdpi);
        if (mr > 0) { autoCache.put(key, mr); return mr; }
        long p = Math.round(sizePt * scale * vdpi / 72.0);
        String n = FontManager.norm(font == null ? fontName : font);
        if (isMs(n)) { autoCache.put(key, msRow(p)); return msRow(p); }
        FontManager.Spec sp = fm.spec(font, false, false);
        Double re = ROW_EM.get(n);
        // 標準フォントの行 (rowPx) と同じく、実測表のある書体はその高さ (表を見ないと同じフォントでも 2 つの計算が食い違う)
        long h = re != null ? Math.round(re * p) : Math.round(sp.ascent * p) + Math.round(sp.descent * p);
        int r = (int) (h + (h + 4) / 10 + 1);
        autoCache.put(key, r);
        return r;
    }

    /** 標準フォント名。 */
    public String defaultFont() { return fontName; }

    /** 標準フォントのサイズ (pt)。 */
    public double defaultSize() { return size; }

    /** 印刷倍率ごとのレイアウト係数。 */
    public static final class Factors {
        /** プリンタの最大数字幅 (デバイスピクセル。端数あり)。 */
        public double mdwP;
        /** 行の換算比 (画面 1 px あたりのデバイスピクセル)。 */
        public double rowRatio;
        /** デバイス 1 ピクセルあたりのシートの pt (横・縦)。 */
        public double ptX, ptY;
    }

    /**
     * 印刷倍率と解像度に対するレイアウト係数を求める。
     *
     * @param scale 印刷倍率 (1.0 = 100%)
     * @param hdpi  横の解像度
     * @param vdpi  縦の解像度
     */
    public Factors factors(double scale, int hdpi, int vdpi) {
        Factors f = new Factors();
        // 実体化する ppem は四捨五入 (偶数丸め・切り上げは不採用)
        long px = Math.round(size * scale * hdpi / 72.0);
        long py = Math.round(size * scale * vdpi / 72.0);
        // 古い Excel で作成 (lowestEdited ≤ 6) し Excel 2016 以降で最後に編集 (lastEdited ≥ 7) したブックの ＭＳ 系は、
        // プリンタの数字幅を丸めない (ppem 89 → 44.5 px)。
        boolean legacy = lastEdited >= 7 && lowestEdited > 0 && lowestEdited <= 6;
        f.mdwP = ms && legacy && authorEnv ? digitEm * px : Math.max(1, Math.round(digitEm * px));
        int mr = measuredPrintRow(fontName, size, scale, vdpi);
        double rp = mr > 0 ? mr : rowPx(py);
        // プリンタの標準行を twip (1/1440 インチ) に切り上げる
        f.rowRatio = rp / (double) screenRow;
        f.ptX = 72.0 / hdpi / scale;
        f.ptY = 72.0 / vdpi / scale;
        return f;
    }
}
