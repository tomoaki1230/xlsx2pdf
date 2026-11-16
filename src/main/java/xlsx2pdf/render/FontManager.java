package xlsx2pdf.render;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.fontbox.ttf.OpenTypeFont;
import org.apache.fontbox.ttf.TTFParser;
import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.io.RandomAccessReadBuffer;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDType0Font;

/**
 * Excel のフォント名を、埋め込める TrueType フォントに解決する。
 *
 * <p>システムのフォントフォルダ (または指定のフォルダ) にある実フォントを優先する。なければ jar に同梱した
 * 寸法互換の代替フォント (Calibri には Carlito、ＭＳ 系には IPA フォントなど) を使う。レイアウトの寸法 (行の高さ) は、
 * 分かっていれば常に元の Excel のフォントの値を使うので、代替フォントで描いても縦の位置は Excel と一致する。
 */
public final class FontManager {
    /** 描画に使う具体的なフォント (書体)。 */
    public static final class Face {
        /** 識別キー (システムフォントは "sys:ファイル:PostScript 名"、同梱フォントはファイル名)。 */
        public final String key;
        final TrueTypeFont ttf;
        /** 埋め込んだ PDF フォント (初回使用時に作る)。 */
        PDType0Font pdf;
        final float unitsPerEm;
        /** グリフ元のアセント・ディセント (em、正の値)。 */
        final float ascent, descent;
        /** 書体自体が太字・斜体か (でなければ合成する)。 */
        final boolean bold, italic;

        Face(String key, TrueTypeFont ttf) throws IOException {
            this.key = key;
            this.ttf = ttf;
            int ms = ttf.getHeader() != null ? ttf.getHeader().getMacStyle() : 0;
            String k = key.toLowerCase(Locale.ROOT);
            this.bold = (ms & 1) != 0 || k.contains("bold");
            this.italic = (ms & 2) != 0 || k.contains("italic") || k.contains("oblique");
            this.unitsPerEm = ttf.getUnitsPerEm();
            float a, d;
            if (ttf.getOS2Windows() != null && ttf.getOS2Windows().getWinAscent() > 0) {
                a = ttf.getOS2Windows().getWinAscent();
                d = ttf.getOS2Windows().getWinDescent();
            } else {
                a = ttf.getHorizontalHeader().getAscender();
                d = -ttf.getHorizontalHeader().getDescender();
            }
            this.ascent = a / unitsPerEm;
            this.descent = d / unitsPerEm;
            var os2 = ttf.getOS2Windows();
            this.useTypo = os2 != null && (os2.getFsSelection() & 128) != 0 && os2.getTypoAscender() > 0;
            this.typoAscent = useTypo ? os2.getTypoAscender() / (float) unitsPerEm : ascent;
            this.typoDescent = useTypo ? -os2.getTypoDescender() / (float) unitsPerEm : descent;
        }

        /** OS/2 の fsSelection の USE_TYPO_METRICS (bit 7) が立っているか。図形の文字の行の高さは typo の値 (Aptos など)。 */
        public final boolean useTypo;
        /** 図形の文字の行のアセント・ディセント (em。USE_TYPO_METRICS なら typo、なければ win の値)。 */
        public final float typoAscent, typoDescent;

        /** コードポイントのグリフを持つか。 */
        public boolean hasGlyph(int cp) {
            if (ComplexText.isCluster(cp)) cp = ComplexText.firstCp(cp);
            try {
                return ttf.getUnicodeCmapLookup().getGlyphId(cp) > 0;
            } catch (IOException e) {
                return false;
            }
        }

        /** 送り幅 (em)。 */
        public float advance(int cp) {
            if (ComplexText.isCluster(cp)) return shaped(cp).adv;
            try {
                int gid = ttf.getUnicodeCmapLookup().getGlyphId(cp);
                return ttf.getAdvanceWidth(gid) / unitsPerEm;
            } catch (IOException e) {
                return 0.5f;
            }
        }
    
        /** AWT の書体の元 (複雑な文字体系の整形に使う。{@link ComplexText})。システムフォントはファイルと PostScript 名、同梱フォントは中身。 */
        java.io.File awtFile;
        String awtPs;
        byte[] awtData;
        private java.awt.Font awt;
        private boolean awtTried;
        private final Map<Integer, ComplexText.Shaped> shaped = new java.util.concurrent.ConcurrentHashMap<>();

        /** 整形に使う AWT の書体 (作れなければ null)。 */
        synchronized java.awt.Font awtFont() {
            if (awtTried) return awt;
            awtTried = true;
            try {
                if (awtFile != null) {
                    java.awt.Font[] fs = java.awt.Font.createFonts(awtFile);
                    for (java.awt.Font f : fs) if (awtPs == null || awtPs.equals(f.getPSName())) { awt = f; break; }
                    if (awt == null && fs.length > 0) awt = fs[0];
                } else if (awtData != null) {
                    awt = java.awt.Font.createFont(java.awt.Font.TRUETYPE_FONT, new java.io.ByteArrayInputStream(awtData));
                }
            } catch (java.awt.FontFormatException | IOException | RuntimeException e) {
                awt = null;
            }
            return awt;
        }

        /** 複雑な文字体系のまとまりを整形した結果 (送り幅と輪郭)。 */
        public ComplexText.Shaped shaped(int cp) {
            return shaped.computeIfAbsent(cp, k -> ComplexText.shape(this, k));
        }
    }

    /** Excel のフォント名とスタイルに対して選んだフォント。 */
    public static final class Spec {
        /** 主なグリフ元。 */
        public final Face face;
        /** CJK 文字のグリフ元。 */
        public final Face eaFace;
        /** 太字・斜体の指定。 */
        public final boolean bold;
        public final boolean italic;
        /** 元のフォントのアセント・ディセント (em)。 */
        public final float ascent;
        public final float descent;
        /** 要求されたフォント名。 */
        public final String requested;
        /**
         * 図形の文字で、CJK の文字のない行に使うアセント・ディセント (em)。欧文の書体だけの値 (和文の書体を混ぜない。
         * USE_TYPO_METRICS の書体は typo の値。Aptos Narrow 11pt の行送りは 13.44pt)。
         */
        public float textAscent, textDescent;
        /**
         * 図形の文字の 1 行の基線より上・下 (em。20 書体 × 12・20pt)。和文の書体は高さが
         * (winAscent + winDescent) × 1.3 で、余りの 0.3 は上下に半分ずつ。欧文の書体は hhea のアセント + lineGap が上、ディセントが下。
         */
        public float lineAsc, lineDesc;
        /** 和文の書体か (ひらがなを持つ)。 */
        public boolean eaFont;

        Spec(Face face, Face eaFace, boolean bold, boolean italic, float ascent, float descent, String requested) {
            this.face = face;
            this.eaFace = eaFace;
            this.bold = bold;
            this.italic = italic;
            this.ascent = ascent;
            this.descent = descent;
            this.requested = requested;
            this.textAscent = ascent;
            this.textDescent = descent;
            this.lineAsc = ascent;
            this.lineDesc = descent;
        }

        /** 行の高さ (em)。 */
        public float lineHeight() { return ascent + descent; }

        /** 書体 f で太字を合成するか。 */
        public boolean synthBold(Face f) { return bold && !f.bold; }

        /** 書体 f で斜体を合成するか。 */
        public boolean synthItalic(Face f) { return italic && !f.italic; }

        /**
         * 字形の輪郭 (基線の原点、y は下向き、pt)。送り幅を adv[0] に入れる。輪郭が取れなければ null。
         */
        public java.awt.geom.Path2D outline(int cp, float size, FontManager fm, double[] adv) {
            Face f = faceFor(cp, fm);
            if (ComplexText.isCluster(cp)) {
                ComplexText.Shaped sh = f.shaped(cp);
                adv[0] = sh.adv * size;
                return new java.awt.geom.Path2D.Double(sh.path, java.awt.geom.AffineTransform.getScaleInstance(size, size));
            }
            try {
                var cmap = f.ttf.getUnicodeCmapLookup();
                int gid = cmap == null ? 0 : cmap.getGlyphId(cp);
                double k = size / f.unitsPerEm;
                adv[0] = f.ttf.getAdvanceWidth(gid) * k;
                var gd = f.ttf.getGlyph() == null ? null : f.ttf.getGlyph().getGlyph(gid);
                if (gd == null) return new java.awt.geom.Path2D.Double();
                java.awt.geom.AffineTransform at = new java.awt.geom.AffineTransform(k, 0, 0, -k, 0, 0);
                return new java.awt.geom.Path2D.Double(gd.getPath(), at);
            } catch (IOException | RuntimeException e) {
                return null;
            }
        }

        /** コードポイントを描く書体 (主 → CJK → 予備の順。どれにもなければ主)。 */
        public Face faceFor(int cp, FontManager fm) {
            if (face.hasGlyph(cp) && !(isCjk(cp) && face != eaFace && eaFace.hasGlyph(cp) && preferEa(face))) return face;
            if (eaFace.hasGlyph(cp)) return eaFace;
            for (Face f : fm.fallbacks()) if (f.hasGlyph(cp)) return f;
            return face;
        }

        private static boolean preferEa(Face f) {
            // Liberation など欧文の書体が一部の CJK 約物を持つことがある。Excel はそれらを東アジアのフォントから
            // 取るので、同じようにする
            return f.key.startsWith("Liberation") || f.key.startsWith("Carlito") || f.key.startsWith("Caladea");
        }

        /** {@link #msGrid()} の結果のキャッシュ。 */
        private Boolean msGrid;

        /** 等幅の ＭＳ ゴシック/明朝 (半角のピクセル格子でヒンティングされる。実測: プロポーショナル書体は違う)。 */
        boolean msGrid() {
            if (msGrid == null) {
                String n = requested == null ? "" : norm(requested);
                msGrid = PrinterModel.isMs(n) && !((n.startsWith("ms p") || n.startsWith("msp")));
            }
            return msGrid;
        }

        /** 文字列の幅 (pt)。 */
        public float width(String s, float size, FontManager fm) {
            s = ComplexText.cluster(s);
            float w = 0;
            for (int i = 0; i < s.length(); ) {
                int cp = s.codePointAt(i);
                i += Character.charCount(cp);
                w += charWidth(cp, size, fm);
            }
            return w;
        }

        /** 1 文字の送り幅 (pt。合成太字とピクセル格子への吸着を含む)。 */
        public float charWidth(int cp, float size, FontManager fm) {
            if (cp == '\n' || cp == '\r') return 0;
            // 数字幅の空白 (分数の ? の空き): そのフォントの数字の幅 (U+2007 の字形がないフォントでも)
            if (cp == 0x2007) return charWidth('0', size, fm);
            if (isPad(cp)) return charWidth(cp - PAD_BASE, size, fm);
            Face f = faceFor(cp, fm);
            float adv = f.advance(cp);
            boolean synth = synthBold(f);
            double[] g = GRID.get();
            // 太字のない書体の合成太字: GDI は送り幅を 1 デバイスピクセル広げる
            boolean onePx = synth && g != null;
            if (synth && !onePx) adv += 0.02f;
            double extra = onePx ? g[0] : 0;
            if (g != null && msGrid()) {
                // ＭＳ ゴシック/明朝は、全角がちょうど半角 2 つ分になるようヒンティングされる:
                // 送り幅は round(ppem / 2) デバイスピクセルの倍数 (ppem 61 → 半角 31、全角 62)
                long ppem = Math.round(size / g[0]);
                long half = (ppem + 1) / 2;
                return (float) (Math.round(adv * 2 * half) * g[0] + extra);
            }
            return (float) (snap(adv * size) + extra);
        }
    }

    /** 表示形式の「_x」の空白を表す私用の文字の始まり (U+E000 + x。描かない)。 */
    public static final int PAD_BASE = 0xE000;

    /** 表示形式の「_x」の空白か。 */
    public static boolean isPad(int cp) {
        return cp >= PAD_BASE && cp < PAD_BASE + 0x100;
    }

    // ---- よく使う Excel のフォントの元の寸法 (winAscent/winDescent、em) ----
    private static final Map<String, float[]> METRICS = new HashMap<>();

    static {
        m(0.952f, 0.269f, "calibri", "calibri light");
        m(0.905f, 0.212f, "arial", "arial narrow", "helvetica");
        m(0.891f, 0.216f, "times new roman", "times");
        m(0.833f, 0.300f, "courier new", "courier");
        m(0.950f, 0.222f, "cambria");
        m(1.000f, 0.210f, "tahoma");
        m(1.005f, 0.210f, "verdana");
        m(1.050f, 0.250f, "segoe ui");
        m(1.007f, 0.222f, "century");
        m(0.859f, 0.141f, "ｍｓ ｐゴシック", "ms pgothic", "ｍｓ ゴシック", "ms gothic", "ｍｓ 明朝", "ms mincho",
                "ｍｓ ｐ明朝", "ms pmincho", "ms ui gothic", "ｍｓ ｕｉ ゴシック");
        m(1.07f, 0.43f, "游ゴシック", "yu gothic", "游ゴシック light", "yu gothic light", "游ゴシック medium",
                "yu gothic medium", "yu gothic ui");
        m(0.880f, 0.240f, "游明朝", "yu mincho", "游明朝 demibold", "游明朝 light");
        m(1.075f, 0.425f, "メイリオ", "meiryo");
        m(0.88f, 0.25f, "meiryo ui"); // 11pt の標準行 15.75pt (21px)
        m(0.880f, 0.120f, "hg丸ｺﾞｼｯｸm-pro", "hgｺﾞｼｯｸm", "hgｺﾞｼｯｸe", "hgp創英角ｺﾞｼｯｸub", "hg創英角ｺﾞｼｯｸub",
                "hg明朝e", "hgp明朝e", "hg明朝b", "hgs明朝e", "hg正楷書体-pro", "hgp教科書体", "hg教科書体");
        m(0.880f, 0.120f, "bizud明朝", "biz ud明朝", "bizudゴシック", "biz udゴシック", "bizudpゴシック",
                "biz udpゴシック", "biz udp明朝", "bizudp明朝", "biz udp明朝 medium", "biz ud明朝 medium");
        // Excel の実効値
        m(0.860f, 0.140f, "hg丸ｺﾞｼｯｸm-pro", "hgｺﾞｼｯｸm", "hgｺﾞｼｯｸe", "hgp創英角ｺﾞｼｯｸub", "hg創英角ｺﾞｼｯｸub", "hg明朝e", "hgp明朝e", "hg明朝b", "hgs明朝e", "hg正楷書体-pro", "hgp教科書体", "hg教科書体");
        m(1.120f, 0.445f, "游ゴシック", "yu gothic", "游ゴシック light", "yu gothic light", "游ゴシック medium", "yu gothic medium", "yu gothic ui");
        m(1.130f, 0.445f, "游明朝", "yu mincho", "游明朝 demibold", "游明朝 light");
        m(1.045f, 0.500f, "メイリオ", "meiryo");
        m(1.045f, 0.270f, "meiryo ui");
        m(0.935f, 0.267f, "calibri", "calibri light");
        m(0.915f, 0.220f, "arial", "arial narrow", "helvetica");
        m(0.967f, 0.213f, "century");
        m(0.915f, 0.245f, "times new roman", "times");
        m(0.930f, 0.220f, "cambria");
        m(0.985f, 0.208f, "tahoma", "verdana");
        m(1.055f, 0.330f, "segoe ui");
        m(0.990f, 0.280f, "aptos", "aptos narrow", "aptos display");
        m(0.980f, 0.250f, "seaford");
        m(0.860f, 0.270f, "ud デジタル 教科書体 n", "ud デジタル 教科書体 nk", "ud デジタル 教科書体 np");
    }

    // ---- グリフの送り幅のデバイスピクセル格子 (Excel のセルの文字は GDI がプリンタのピクセルでレイアウトする) ----
    /** 現在の格子の単位 {シートの pt / プリンタのピクセル} (スレッドごと。null なら無効)。 */
    private static final ThreadLocal<double[]> GRID = new ThreadLocal<>();

    /** {@link #clearGrid()} までグリフの送り幅を unit (シートの pt / プリンタのピクセル) の倍数に丸める。0 なら無効。 */
    public static void setGrid(double unit) {
        if (unit <= 0) GRID.remove(); else GRID.set(new double[] {unit});
    }

    /** 格子への吸着を終える。 */
    public static void clearGrid() { GRID.remove(); }

    /** 今の格子の単位 (吸着していなければ 0)。入れ子で設定するときに、元の値を戻すために使う。 */
    public static double gridUnit() {
        double[] g = GRID.get();
        return g == null ? 0 : g[0];
    }

    /** 送り幅を現在のデバイスピクセル格子に丸める (GDI はグリフごとに送り幅を整数ピクセルに丸める)。 */
    public static float snap(float w) {
        double[] g = GRID.get();
        if (g == null || w == 0) return w;
        return (float) (Math.round(w / g[0]) * g[0]);
    }

    /** 格子への吸着が有効か。 */
    public static boolean gridActive() { return GRID.get() != null; }

    // ---- 実際の Windows フォントの GDI 寸法 (リソース fontmetrics.tsv) ----
    /** 正規化したフォント名 → {unitsPerEm, winAscent, winDescent, 最大数字送り幅}。 */
    private static final Map<String, int[]> GDI = new HashMap<>();

    static {
        try (InputStream in = FontManager.class.getResourceAsStream("/fontmetrics.tsv")) {
            if (in != null) {
                for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\n")) {
                    String[] p = l.split("\t");
                    if (p.length < 5) continue;
                    GDI.put(p[0], new int[] {Integer.parseInt(p[1]), Integer.parseInt(p[2]), Integer.parseInt(p[3]), Integer.parseInt(p[4])});
                }
            }
        } catch (IOException | RuntimeException ignored) {
            // 表はなくてもよい
        }
    }

    /**
     * GDI から見た実際の Windows フォントの寸法 (em): {winAscent, winDescent, 最大数字送り幅}。表になければ null。
     * Excel は、どのフォントでも同じ式で、ここから列幅 (最大数字幅) と行高 (アセント + ディセント) を求める。
     */
    public static double[] gdiMetrics(String name) {
        if (name == null) return null;
        String n = norm(name);
        int[] v = lookupLoose(GDI, n);
        if (v == null) return null;
        double u = v[0];
        Integer hinted = HINTED_DIGIT.get(n);
        return new double[] {v[1] / u, v[2] / u, (hinted != null ? hinted : v[3]) / u};
    }

    /** GDI の表のアウトライン (ヒンティング前) の最大数字送り幅 (em)。なければ -1。 */
    public static double outlineDigit(String name) {
        if (name == null) return -1;
        int[] v = lookupLoose(GDI, norm(name));
        return v == null ? -1 : v[3] / (double) v[0];
    }

    /**
     * TrueType のヒンティング後の数字送り幅 (アウトラインの値と違うもの、2048 単位)。Excel の PDF で実測:
     * 游ゴシック 11pt は 600 dpi/100% で 51 デバイス px、83% で 43 (アウトライン 1139/2048 では 51/42)。
     */
    private static final Map<String, Integer> HINTED_DIGIT = Map.of("游ゴシック", 1146, "yu gothic", 1146);

    /** 元の寸法の表に登録する。 */
    private static void m(float a, float d, String... names) {
        for (String n : names) METRICS.put(n, new float[] {a, d});
    }

    private final PDDocument doc;
    /** 読み込んだ書体 (キー → 書体)。 */
    private final Map<String, Face> faces = new HashMap<>();
    /** 選んだフォント (名前|太字|斜体 → フォント)。 */
    private final Map<String, Spec> specs = new HashMap<>();
    /** システムフォント: 正規化したファミリ名 → スタイル (bit0 太字、bit1 斜体) → 場所。 */
    private final Map<String, Map<Integer, Loc>> systemFonts = new HashMap<>();
    /** どのフォントにもないグリフの予備の書体。 */
    private List<Face> fallbackFaces;

    /**
     * フォントファイル内の書体の場所。
     *
     * @param psName     PostScript 名
     * @param collection TrueType コレクション (.ttc) か
     */
    private record Loc(File file, String psName, boolean collection) {}

    /**
     * フォントを検索する準備をする (システムと指定のフォルダを走査する)。
     *
     * @param doc       フォントを埋め込む PDF
     * @param extraDirs 追加で検索するフォルダ (システムより優先)
     */
    public FontManager(PDDocument doc, List<File> extraDirs) {
        this(doc, extraDirs, defaultCacheDir());
    }

    /** フォントの走査結果のキャッシュの既定のフォルダ (~/.cache/xlsx2pdf)。 */
    public static File defaultCacheDir() {
        return new File(System.getProperty("user.home"), ".cache/xlsx2pdf");
    }

    /**
     * フォントを検索する準備をする (システムと指定のフォルダを走査する)。
     *
     * @param doc       フォントを埋め込む PDF
     * @param extraDirs 追加で検索するフォルダ (システムより優先)
     * @param cacheDir  走査結果のキャッシュ (fonts-*.tsv) を置くフォルダ。null ならキャッシュを読まず、書かない
     */
    public FontManager(PDDocument doc, List<File> extraDirs, File cacheDir) {
        this.doc = doc;
        List<File> dirs = new ArrayList<>(extraDirs);
        String home = System.getProperty("user.home");
        dirs.add(new File("C:\\Windows\\Fonts"));
        dirs.add(new File("/mnt/c/Windows/Fonts")); // WSL
        dirs.add(new File(home, "AppData/Local/Microsoft/Windows/Fonts"));
        // Office のクラウドフォント
        dirs.add(new File(home, "AppData/Local/Microsoft/FontCache/4/CloudFonts"));
        File[] wslUsers = new File("/mnt/c/Users").listFiles(File::isDirectory);   // WSL: Windows のユーザーのフォント
        if (wslUsers != null)
            for (File u : wslUsers) {
                dirs.add(new File(u, "AppData/Local/Microsoft/Windows/Fonts"));
                dirs.add(new File(u, "AppData/Local/Microsoft/FontCache/4/CloudFonts"));
            }
        // Office に付属するフォント (HG 系など。Excel はここのフォントも使う)
        dirs.add(new File("C:\\Program Files\\Microsoft Office\\root\\vfs\\Fonts\\private"));
        dirs.add(new File("/mnt/c/Program Files/Microsoft Office/root/vfs/Fonts/private"));
        dirs.add(new File("/Library/Fonts"));
        dirs.add(new File(home, "Library/Fonts"));
        dirs.add(new File("/usr/share/fonts"));
        dirs.add(new File("/usr/local/share/fonts"));
        dirs.add(new File(home, ".fonts"));
        dirs.add(new File(home, ".local/share/fonts"));
        for (File d : dirs) {
            Map<String, Map<Integer, Loc>> found = SCAN_CACHE.computeIfAbsent(d.getAbsolutePath(), k -> scanCached(d, cacheDir));
            for (var e : found.entrySet()) {
                Map<Integer, Loc> t = systemFonts.computeIfAbsent(e.getKey(), x -> new HashMap<>());
                for (var st : e.getValue().entrySet()) t.putIfAbsent(st.getKey(), st.getValue());
            }
        }
    }

    /** AWT のフォント (EMF の描画用) のキャッシュ (ファイル:PostScript 名 → フォント)。 */
    private static final Map<String, java.awt.Font> AWT_FONTS = new java.util.concurrent.ConcurrentHashMap<>();

    /**
     * Excel のフォント名に対する AWT のフォント (大きさ 1pt) を返す (EMF のラスタライズ用)。実フォントがなければ同梱の代替。
     *
     * @param name フォント名 (先頭の @ (縦書き) は除いて渡す)
     */
    public java.awt.Font awtFont(String name, boolean bold, boolean italic) {
        int style = (bold ? 1 : 0) | (italic ? 2 : 0);
        Map<Integer, Loc> m = systemFonts.get(norm(name == null ? "" : name));
        Loc loc = m == null ? null : m.getOrDefault(style, m.get(0));
        try {
            if (loc != null) {
                String key = loc.file + ":" + loc.psName;
                java.awt.Font f = AWT_FONTS.get(key);
                if (f == null) {
                    java.awt.Font[] all = java.awt.Font.createFonts(loc.file);
                    f = all[0];
                    for (java.awt.Font x : all) if (x.getPSName().equalsIgnoreCase(loc.psName)) f = x;
                    AWT_FONTS.put(key, f);
                }
                return f;
            }
            String sub = substitute(name == null ? "" : name);
            String file = sub.startsWith("ipa") ? sub + ".ttf" : sub + "-" + STYLE_SUFFIX[style] + ".ttf";
            java.awt.Font f = AWT_FONTS.get("bundled:" + file);
            if (f == null) {
                try (InputStream in = FontManager.class.getResourceAsStream("/fonts/" + file)) {
                    f = java.awt.Font.createFont(java.awt.Font.TRUETYPE_FONT, in);
                }
                AWT_FONTS.put("bundled:" + file, f);
            }
            return f;
        } catch (Exception e) {
            return new java.awt.Font(java.awt.Font.SANS_SERIF, style, 1);
        }
    }

    /**
     * PostScript 名 (「MS-PGothic」など。サブセットの接頭辞と ",Bold" などは除く) でシステムのフォントを探す
     * 見つからなければ null。
     */
    public static TrueTypeFont trueTypeByPsName(String psName) {
        if (psName == null) return null;
        String want = psName.replaceFirst("^[A-Z]{6}\\+", "").replaceFirst(",.*$", "").replace("-", "").replace(" ", "").toLowerCase(Locale.ROOT);
        for (Map<String, Map<Integer, Loc>> dir : SCAN_CACHE.values()) {
            for (Map<Integer, Loc> m : dir.values()) {
                for (Loc loc : m.values()) {
                    if (loc.psName == null || !loc.psName.replace("-", "").replace(" ", "").toLowerCase(Locale.ROOT).equals(want)) continue;
                    try {
                        if (loc.collection) {
                            TrueTypeCollection c = new TrueTypeCollection(new java.io.ByteArrayInputStream(java.nio.file.Files.readAllBytes(loc.file.toPath())));
                            TrueTypeFont t = c.getFontByName(loc.psName);
                            if (t != null) return t;
                        } else {
                            return new TTFParser().parse(new RandomAccessReadBuffer(java.nio.file.Files.readAllBytes(loc.file.toPath())));
                        }
                    } catch (IOException e) {
                        return null;
                    }
                }
            }
        }
        return null;
    }

    /** {@link #withRegular} に渡す処理。 */
    interface FontReader<T> { T read(TrueTypeFont t) throws IOException; }

    /** システムの標準の書体 (太字・斜体でない) を開いて読む。見つからなければ null。 */
    static <T> T withRegular(String name, FontReader<T> reader) throws IOException {
        String n = norm(name);
        for (Map<String, Map<Integer, Loc>> dir : SCAN_CACHE.values()) {
            Map<Integer, Loc> m = lookupLoose(dir, n);
            Loc loc = m != null ? m.get(0) : null;
            if (loc == null) continue;
            if (loc.collection) {
                try (TrueTypeCollection c = new TrueTypeCollection(loc.file)) {
                    TrueTypeFont t = c.getFontByName(loc.psName);
                    return t != null ? reader.read(t) : null;
                }
            }
            try (TrueTypeFont t = new TTFParser().parse(new RandomAccessReadBuffer(java.nio.file.Files.readAllBytes(loc.file.toPath())))) {
                return reader.read(t);
            }
        }
        return null;
    }

    /** フォントフォルダの走査結果は、同じプロセス内のすべての変換で共有する (一括比較用)。 */
    private static final Map<String, Map<String, Map<Integer, Loc>>> SCAN_CACHE = new java.util.concurrent.ConcurrentHashMap<>();

    /**
     * フォントフォルダを走査する。大きさと更新日時が変わっていないファイルは cacheDir/fonts-*.tsv を再利用する (cacheDir が null なら使わない)
     * (WSL の /mnt/c のような遅いマウントですべてのフォントを読むと数十秒かかる)。
     */
    private static Map<String, Map<Integer, Loc>> scanCached(File dir, File cacheDir) {
        Map<String, Map<Integer, Loc>> out = new HashMap<>();
        if (!dir.isDirectory()) return out;
        File cache = cacheDir == null ? null
                : new File(cacheDir, "fonts-" + Integer.toHexString(dir.getAbsolutePath().hashCode()) + ".tsv");
        // ファイル → "大きさ\t更新日時" と、その中の書体 (ファミリ、スタイル、PostScript 名、コレクションか)
        Map<String, String> stamp = new HashMap<>();
        Map<String, List<String[]>> entries = new HashMap<>();
        try {
            if (cache != null && cache.isFile()) {
                for (String l : java.nio.file.Files.readAllLines(cache.toPath(), java.nio.charset.StandardCharsets.UTF_8)) {
                    String[] p = l.split("\t", -1);
                    if (p.length < 7) continue;
                    stamp.put(p[0], p[1] + "\t" + p[2]);
                    entries.computeIfAbsent(p[0], x -> new ArrayList<>()).add(new String[] {p[3], p[4], p[5], p[6]});
                }
            }
        } catch (IOException | RuntimeException e) {
            stamp.clear();
            entries.clear();
        }
        List<File> files = new ArrayList<>();
        listFonts(dir, 0, files);
        StringBuilder sb = new StringBuilder();
        boolean changed = false;
        for (File f : files) {
            String path = f.getAbsolutePath(), st = f.length() + "\t" + f.lastModified();
            List<String[]> es = st.equals(stamp.get(path)) ? entries.get(path) : null;
            if (es == null) {
                changed = true;
                Map<String, Map<Integer, Loc>> one = new HashMap<>();
                readFont(f, one);
                es = new ArrayList<>();
                for (var e : one.entrySet())
                    for (var x : e.getValue().entrySet())
                        es.add(new String[] {e.getKey(), String.valueOf(x.getKey()), x.getValue().psName(), String.valueOf(x.getValue().collection())});
                if (es.isEmpty()) es.add(new String[] {"", "-1", "", "false"}); // 使えないファイルも記録する
            }
            for (String[] e : es) {
                sb.append(path).append('\t').append(st).append('\t').append(String.join("\t", e)).append('\n');
                int style = Integer.parseInt(e[1]);
                if (style >= 0) out.computeIfAbsent(e[0], x -> new HashMap<>()).putIfAbsent(style, new Loc(f, e[2], Boolean.parseBoolean(e[3])));
            }
        }
        if (cache != null && (changed || files.size() != stamp.size())) {
            try {
                cache.getParentFile().mkdirs();
                File tmp = new File(cache.getPath() + "." + ProcessHandle.current().pid());
                java.nio.file.Files.writeString(tmp.toPath(), sb, java.nio.charset.StandardCharsets.UTF_8);
                java.nio.file.Files.move(tmp.toPath(), cache.toPath(), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            } catch (IOException | RuntimeException ignored) {
                // キャッシュはなくてもよい
            }
        }
        return out;
    }

    /** フォルダ以下の .ttf / .ttc を列挙する (深さ 5 まで)。 */
    private static void listFonts(File dir, int depth, List<File> out) {
        File[] list = dir.listFiles();
        if (list == null || depth > 5) return;
        java.util.Arrays.sort(list);
        for (File f : list) {
            if (f.isDirectory()) { listFonts(f, depth + 1, out); continue; }
            String n = f.getName().toLowerCase(Locale.ROOT);
            if (n.endsWith(".ttf") || n.endsWith(".ttc")) out.add(f);
        }
    }

    /** フォントファイルを読んで、書体を登録する。 */
    private static void readFont(File f, Map<String, Map<Integer, Loc>> out) {
        String n = f.getName().toLowerCase(Locale.ROOT);
        try {
            if (n.endsWith(".ttf")) {
                TrueTypeFont t = new TTFParser(true).parse(new RandomAccessReadBuffer(java.nio.file.Files.readAllBytes(f.toPath())));
                register(t, f, false, out);
                t.close();
            } else if (n.endsWith(".ttc")) {
                try (TrueTypeCollection c = new TrueTypeCollection(f)) {
                    c.processAllFonts(t -> register(t, f, true, out));
                }
            }
        } catch (Exception | Error e) {
            // 読めないフォントは無視する
        }
    }

    /** 書体を、ファミリ名と完全名 (name ID 1, 4) で登録する。 */
    private static void register(TrueTypeFont t, File f, boolean collection, Map<String, Map<Integer, Loc>> out) throws IOException {
        if (t instanceof OpenTypeFont otf && otf.isPostScript()) return; // CFF のアウトラインは埋め込めない
        if (t.getNaming() == null) return;
        String ps = t.getName();
        int style = 0;
        if (t.getHeader() != null) style = t.getHeader().getMacStyle() & 3;
        List<String> names = new ArrayList<>();
        for (var rec : t.getNaming().getNameRecords()) {
            if (rec.getNameId() == 1 || rec.getNameId() == 4) names.add(rec.getString());
        }
        for (String fam : names) {
            if (fam == null) continue;
            String k = norm(fam);
            out.computeIfAbsent(k, x -> new HashMap<>()).putIfAbsent(style, new Loc(f, ps, collection));
        }
    }

    /** 空白を除いた名前でも照合する表の検索 ("MSPゴシック" = "ＭＳ Ｐゴシック")。 */
    static <T> T lookupLoose(Map<String, T> m, String n) {
        T v = m.get(n);
        if (v != null || n == null) return v;
        String k = n.replace(" ", "");
        for (Map.Entry<String, T> e : m.entrySet()) if (e.getKey().replace(" ", "").equals(k)) return e.getValue();
        return null;
    }

    /** フォント名を正規化する: 小文字、全角英数を半角、空白をまとめる。古い和文フォント名は別名に置き換える。 */
    public static String norm(String s) {
        StringBuilder sb = new StringBuilder();
        for (char c : s.toCharArray()) {
            if (c >= '\uFF01' && c <= '\uFF5E') c = (char) (c - 0xFEE0);
            if (c == '\u3000') c = ' ';
            sb.append(c);
        }
        String n = sb.toString().toLowerCase(Locale.ROOT).replaceAll("\\s+", " ").trim();
        // Windows が置き換える古い和文フォント名 (FontSubstitutes): 明朝 → ＭＳ Ｐ明朝、ゴシック → ＭＳ ゴシック
        String alias = LEGACY_ALIAS.get(n);
        if (alias == null && OSAKA_ALIAS != null && n.equals("osaka")) alias = OSAKA_ALIAS;
        return alias != null ? alias : n;
    }

    /** 別名の置き換えをしない正規化 (小文字、全角英数を半角、空白をまとめる)。 */
    static String rawNorm(String s) {
        StringBuilder sb = new StringBuilder();
        for (char c : s.toCharArray()) {
            if (c >= '\uFF01' && c <= '\uFF5E') c = (char) (c - 0xFEE0);
            if (c == '\u3000') c = ' ';
            sb.append(c);
        }
        return sb.toString().toLowerCase(Locale.ROOT).replaceAll("\\s+", " ").trim();
    }

    /**
     * Windows (Excel の環境) にある書体名か。fontmetrics.tsv・screenrows.tsv にある名前、Windows の FontSubstitutes の名前、
     * フルネーム (ファミリー名 + スタイル: 「UD デジタル 教科書体 NK-R」など) をあるとみなす。
     */
    public static boolean installed(String name) {
        if (name == null) return false;
        String n = rawNorm(name);
        if (n.startsWith("@")) n = n.substring(1);
        if (GDI.containsKey(n) || WIN_SUBST.contains(n) || WIN_FONTS.contains(n) || CLOUD_FONTS.contains(n)) return true;
        for (String k : GDI.keySet())
            if (n.length() > k.length() + 1 && n.startsWith(k) && (n.charAt(k.length()) == '-' || n.charAt(k.length()) == ' ')) return true;
        return false;
    }

    /**
     * ファミリー名ではなく、スタイル付きの名前 (「UD デジタル 教科書体 NK-R」など) でだけ見つかる書体名か。
     * Excel は、この形の名前を charset の指定がなく family の指定があるときには見つけられない (family 1・charset なしは
     * Times New Roman + ＭＳ 明朝、charset 128 か family なしなら UD デジタル 教科書体 NK-R)。
     */
    public static boolean fullNameOnly(String name) {
        if (name == null) return false;
        String n = rawNorm(name);
        if (n.startsWith("@")) n = n.substring(1);
        if (GDI.containsKey(n) || WIN_SUBST.contains(n) || WIN_FONTS.contains(n) || CLOUD_FONTS.contains(n)) return false;
        return installed(name);
    }

    /** Office のクラウドフォント (Windows にはないが Excel が取得して使う)。 */
    private static final java.util.Set<String> CLOUD_FONTS = java.util.Set.of("aptos", "aptos narrow", "aptos display", "seaford");

    /**
     * Windows + Office にある書体名 (winfonts.txt。日本語名と英語名)。画面の行の表 (screenrows.tsv) には、ない書体を Excel が
     * 置き換えた結果の行 (IPA Pゴシック・ＭＳゴシックなど) も入っているので、判定には使わない。
     */
    private static final java.util.Set<String> WIN_FONTS = new java.util.HashSet<>();
    static {
        try (var in = FontManager.class.getResourceAsStream("/winfonts.txt")) {
            if (in != null)
                for (String l : new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).split("\n"))
                    if (!l.startsWith("#") && !l.isBlank()) WIN_FONTS.add(l.trim());
        } catch (java.io.IOException e) {
            // 表がなければ寸法の表だけで判定する
        }
    }

    /** Windows の FontSubstitutes にある名前 (charset・family によらず置き換わる)。 */
    private static final java.util.Set<String> WIN_SUBST = java.util.Set.of("ゴシック", "helvetica", "helv", "tms rmn",
            "times", "courier", "ms shell dlg", "ms shell dlg 2", "ms sans serif", "ms serif");

    /** Mac の Osaka を Windows で置き換える書体 (日本語の等幅系 → ＭＳ ゴシック)。 */
    static final String OSAKA_ALIAS = "ms ゴシック";

    /** 古い和文フォント名の別名と、空白の抜けた表記。 */
    private static final Map<String, String> LEGACY_ALIAS = Map.ofEntries(Map.entry("明朝", "ms p明朝"), Map.entry("ゴシック", "ms ゴシック"),
            Map.entry("mspゴシック", "ms pゴシック"), Map.entry("msゴシック", "ms ゴシック"), Map.entry("msp明朝", "ms p明朝"),
            Map.entry("ms明朝", "ms 明朝"), Map.entry("helvetica", "arial"), Map.entry("helv", "arial"));

    /** 元のフォントの寸法 {アセント, ディセント} (em)。表になければ null。 */
    private static float[] metricsFor(String name) {
        String n = name.toLowerCase(Locale.ROOT);
        float[] m = METRICS.get(n);
        if (m == null) m = METRICS.get(norm(name));
        if (m == null) {
            // 正規化したキーでも探す
            for (var e : METRICS.entrySet()) if (norm(e.getKey()).equals(norm(name))) return e.getValue();
        }
        return m;
    }

    /** システムフォントの書体を読み込む (なければ null)。 */
    private Face loadSystem(String family, int style) {
        Map<Integer, Loc> m = systemFonts.get(norm(family));
        if (m == null) return null;
        Loc loc = m.get(style);
        if (loc == null) return null;
        String key = "sys:" + loc.file + ":" + loc.psName;
        Face f = faces.get(key);
        if (f != null) return f;
        try {
            TrueTypeFont t;
            if (loc.collection) {
                TrueTypeCollection c = new TrueTypeCollection(new java.io.ByteArrayInputStream(java.nio.file.Files.readAllBytes(loc.file.toPath())));
                t = c.getFontByName(loc.psName);
            } else {
                t = new TTFParser().parse(new RandomAccessReadBuffer(java.nio.file.Files.readAllBytes(loc.file.toPath())));
            }
            if (t == null) return null;
            f = new Face(key, t);
            f.awtFile = loc.file;
            f.awtPs = loc.psName;
            faces.put(key, f);
            return f;
        } catch (IOException e) {
            return null;
        }
    }

    /**
     * 同梱の代替フォントを読み込む。同梱フォントは別の jar (xlsx2pdf-fonts.jar) にあり、本体の jar と同じフォルダに置くと
     * 使われる (本体の jar のマニフェストの Class-Path)。その jar がなければ、PC に入っている似たフォントで代用する。
     *
     * @throws xlsx2pdf.FontUnavailableException 同梱フォントも代わりのフォントもないとき
     */
    Face bundled(String file) {
        Face f = faces.get(file);
        if (f != null) return f;
        try (InputStream in = FontManager.class.getResourceAsStream("/fonts/" + file)) {
            if (in != null) {
                byte[] data = in.readAllBytes();
                TrueTypeFont t = new TTFParser().parse(new RandomAccessReadBuffer(data));
                f = new Face(file.replace(".ttf", ""), t);
                f.awtData = data;
                faces.put(file, f);
                return f;
            }
        } catch (IOException e) {
            // 読めなければ PC のフォントで代用する
        }
        f = systemSubstitute(file);
        if (f == null) {
            throw new xlsx2pdf.FontUnavailableException("代わりに使うフォントが見つかりません (" + file + ")。"
                    + "xlsx2pdf-fonts.jar を xlsx2pdf.jar と同じフォルダに置くか、--font-dir でフォントのフォルダを指定してください");
        }
        if (!substituteWarned) {
            substituteWarned = true;
            System.err.println("注意: 同梱フォント (xlsx2pdf-fonts.jar) がないため、PC のフォントで代用します。"
                    + "Excel と字幅が変わることがあります");
        }
        faces.put(file, f);
        return f;
    }

    /**
     * 同梱フォントの jar がないとき、代わりのフォントが PC にあるかを最初に確かめる (描画の途中で見つからないと、
     * セル単位の例外処理に紛れて文字が欠けるため)。
     *
     * @throws xlsx2pdf.FontUnavailableException 代わりのフォントがないとき
     */
    public void checkFallbacks() {
        if (FontManager.class.getResource("/fonts/ipag.ttf") != null) return;
        fallbacks();
    }

    /** 同梱フォントがないことを一度だけ知らせたか。 */
    private static volatile boolean substituteWarned;

    /** 同梱フォント (ファイル名) の代わりに PC で探すフォント名 (先にあるものほど優先)。 */
    private static List<String> substituteCandidates(String file) {
        String f = file.toLowerCase(Locale.ROOT);
        if (f.startsWith("ipaexg")) return List.of("ipaexgothic", "ipaexゴシック", "yu gothic", "游ゴシック", "meiryo", "ms pgothic", "noto sans cjk jp", "noto sans jp", "ipapgothic");
        if (f.startsWith("ipaexm")) return List.of("ipaexmincho", "ipaex明朝", "yu mincho", "游明朝", "ms pmincho", "noto serif cjk jp", "noto serif jp", "ipapmincho");
        if (f.startsWith("ipagp")) return List.of("ms pgothic", "ms pゴシック", "ipapgothic", "ipapゴシック", "noto sans cjk jp", "noto sans jp", "yu gothic", "meiryo");
        if (f.startsWith("ipag")) return List.of("ms gothic", "ms ゴシック", "ipagothic", "ipaゴシック", "noto sans mono cjk jp", "noto sans cjk jp", "noto sans jp", "yu gothic");
        if (f.startsWith("ipamp")) return List.of("ms pmincho", "ms p明朝", "ipapmincho", "ipap明朝", "noto serif cjk jp", "noto serif jp", "yu mincho");
        if (f.startsWith("ipam")) return List.of("ms mincho", "ms 明朝", "ipamincho", "ipa明朝", "noto serif cjk jp", "noto serif jp", "yu mincho");
        if (f.startsWith("carlito")) return List.of("calibri", "carlito", "arial", "liberation sans", "dejavu sans");
        if (f.startsWith("caladea")) return List.of("cambria", "caladea", "times new roman", "liberation serif", "dejavu serif");
        if (f.startsWith("liberationserif")) return List.of("times new roman", "liberation serif", "dejavu serif");
        if (f.startsWith("liberationmono")) return List.of("courier new", "liberation mono", "dejavu sans mono");
        if (f.startsWith("dejavusans")) return List.of("dejavu sans", "segoe ui symbol", "arial unicode ms", "ms gothic", "arial");
        return List.of("arial", "liberation sans", "dejavu sans", "helvetica");
    }

    /** 同梱フォントの代わりに、PC に入っている似たフォントを探す (なければ null)。 */
    private Face systemSubstitute(String file) {
        int style = 0;
        for (int i = 1; i < STYLE_SUFFIX.length; i++) if (file.endsWith("-" + STYLE_SUFFIX[i] + ".ttf")) style = i;
        for (String name : substituteCandidates(file)) {
            Face f = loadSystem(name, style);
            if (f == null && style != 0) f = loadSystem(name, 0);
            if (f != null) return f;
        }
        return null;
    }

    /** 予備の書体 (DejaVu は IPA フォントにない ☑ などの記号用)。 */
    List<Face> fallbacks() {
        if (fallbackFaces == null) {
            fallbackFaces = List.of(bundled("ipaexg.ttf"), bundled("LiberationSans-Regular.ttf"), bundled("ipag.ttf"),
                    bundled("DejaVuSans.ttf"));
        }
        return fallbackFaces;
    }

    /** 書体を PDF に埋め込む (サブセット。初回だけ)。 */
    public PDType0Font pdfFont(Face f) throws IOException {
        if (f.pdf == null) {
            // GSUB (合字など) は使わない: Excel はセルの文字に合字を使わず、PDFBox の Latin の GSUB は
            // 文字列ごとに正規表現を作り直すので、文字の多いシートで極端に遅くなる
            f.ttf.setEnableGsub(false);
            f.pdf = PDType0Font.load(doc, f.ttf, true);
        }
        return f.pdf;
    }

    /** 同梱の欧文フォントのファイル名のスタイル (スタイル番号の順)。 */
    private static final String[] STYLE_SUFFIX = {"Regular", "Bold", "Italic", "BoldItalic"};

    /** Excel のフォント名に対する同梱の代替フォント (ファイル名の基部。和文は ipa で始まる)。 */
    static String substitute(String name) {
        String n = norm(name);
        boolean mincho = n.contains("明朝") || n.contains("mincho") || n.contains("教科書") || n.contains("楷書")
                || n.contains("行書") || n.contains("serif");
        boolean gothic = n.contains("ゴシック") || n.contains("ｺﾞｼｯｸ") || n.contains("gothic") || n.contains("メイリオ")
                || n.contains("meiryo") || n.contains("丸") || n.contains("osaka") || n.contains("ヒラギノ");
        boolean cjk = mincho || gothic || n.chars().anyMatch(c -> c > 0x2E80);
        if (cjk) {
            boolean ms = n.startsWith("ms ") || n.startsWith("ｍｓ") || n.startsWith("hg") || n.startsWith("biz")
                    || n.startsWith("ms");
            boolean prop = n.matches("^(ms |ms|hg|biz ?ud)p.*") || n.contains(" p") || n.startsWith("hgp")
                    || n.contains("ud p");
            if (mincho) {
                if (ms) return prop ? "ipamp" : "ipam";
                return "ipaexm";
            }
            if (ms) return prop ? "ipagp" : "ipag";
            if (n.contains("ui")) return "ipagp";
            return "ipaexg";
        }
        if (n.startsWith("calibri") || n.startsWith("candara") || n.startsWith("corbel")) return "Carlito";
        if (n.startsWith("cambria") || n.startsWith("constantia")) return "Caladea";
        if (n.startsWith("times") || n.startsWith("century") || n.startsWith("georgia") || n.startsWith("garamond")
                || n.startsWith("book antiqua") || n.startsWith("palatino") || n.contains("serif")) return "LiberationSerif";
        if (n.startsWith("courier") || n.contains("mono") || n.startsWith("consolas") || n.startsWith("lucida console"))
            return "LiberationMono";
        return "LiberationSans";
    }

    /** 欧文フォントの CJK 文字に使う既定の東アジアフォント (同梱の IPA フォント)。 */
    private String eaDefault = "ipagp";

    /** 既定の東アジアフォントを、Excel のフォント名から決める。 */
    public void setEastAsianDefault(String excelName) {
        if (excelName != null) eaDefault = substitute(excelName);
        if (!eaDefault.startsWith("ipa")) eaDefault = "ipagp";
    }

    /**
     * Excel のフォント名とスタイルに対するフォントを選ぶ (結果はキャッシュする)。
     *
     * @param name フォント名 (空なら Calibri)
     */
    public Spec spec(String name, boolean bold, boolean italic) {
        if (name == null || name.isBlank()) name = "Calibri";
        String key = name + "|" + bold + "|" + italic;
        Spec s = specs.get(key);
        if (s != null) return s;
        int style = (bold ? 1 : 0) | (italic ? 2 : 0);
        Face face = loadSystem(name, style);
        if (face == null) face = loadSystem(name, 0);
        String sub = substitute(name);
        if (face == null && sub.startsWith("ipaex")) {
            // PC にない和文フォント (ＪＳ明朝など) は、Windows (GDI) と同じくＭＳ Ｐ明朝 / ＭＳ Ｐゴシックで代用する
            String win = sub.equals("ipaexm") ? "ms pmincho" : "ms pgothic";
            face = loadSystem(win, style);
            if (face == null && style != 0) face = loadSystem(win, 0);
        }
        // 実フォントがなく同梱の代替で描くとき、太い書体 (HG明朝E、HG創英角ゴシックUB など) は太字を合成する
        if (face == null && heavy(name)) bold = true;
        if (face == null) {
            if (sub.startsWith("ipa")) {
                face = bundled(sub + ".ttf");
            } else {
                face = bundled(sub + "-" + STYLE_SUFFIX[style] + ".ttf");
            }
        }
        // 和文フォントは、実フォントがひらがなを持てばそれを、なければ同梱の代替を CJK に使う
        Face ea = sub.startsWith("ipa") ? (face.hasGlyph(0x3042) ? face : bundled(sub + ".ttf")) : bundled(eaDefault + ".ttf");
        float[] m = metricsFor(name);
        float asc = m != null ? m[0] : face.ascent;
        float desc = m != null ? m[1] : face.descent;
        s = new Spec(face, ea, bold, italic, asc, desc, name);
        if (face.useTypo) { s.textAscent = face.typoAscent; s.textDescent = face.typoDescent; }
        s.eaFont = sub.startsWith("ipa") || face.hasGlyph(0x3042);
        if (s.eaFont) {
            // 実フォントがあれば OS/2 の winAscent・winDescent (游ゴシック 0.985 + 0.302、メイリオ 1.06 + 0.44)。なければ寸法の表の値
            boolean real = face.key.startsWith("sys:") && face.hasGlyph(0x3042);
            float wa = real ? face.ascent : asc, wd = real ? face.descent : desc;
            s.lineAsc = wa + EA_LINE_EXTRA / 2 * (wa + wd);
            s.lineDesc = wd + EA_LINE_EXTRA / 2 * (wa + wd);
        } else if (face.useTypo) {
            s.lineAsc = s.textAscent; s.lineDesc = s.textDescent;
        } else {
            try {
                var hh = face.ttf.getHorizontalHeader();
                s.lineAsc = (hh.getAscender() + hh.getLineGap()) / face.unitsPerEm;
                s.lineDesc = -hh.getDescender() / face.unitsPerEm;
            } catch (IOException | RuntimeException e) {
                s.lineAsc = s.textAscent; s.lineDesc = s.textDescent;
            }
        }
        specs.put(key, s);
        return s;
    }

    /** 図形の文字で、和文の書体の行の高さに足す割合 (winAscent + winDescent に対して)。 */
    static final float EA_LINE_EXTRA = 0.3f;
    /** 欧文の書体の中の和文の文字 (代わりの和文の書体 = ＭＳ Ｐゴシック相当) の基線より上・下 (em)。 */
    static final float EA_FALLBACK_ASC = 0.859f + EA_LINE_EXTRA / 2, EA_FALLBACK_DESC = 0.141f + EA_LINE_EXTRA / 2;

    /** 名前から太い書体と分かるか (HG 系の E・UB・B、極太・太字の表記)。 */
    static boolean heavy(String name) {
        String n = norm(name);
        return n.matches("^hg.*(e|ub|b)$") || n.contains("極太") || n.contains("太ゴ") || n.endsWith(" bold") || n.contains("heavy")
                || n.contains("black");
    }

    /** 東アジアフォントを明示したフォント (DrawingML のランは欧文と東アジアで別の書体を持つ)。 */
    public Spec spec(String latin, String ea, boolean bold, boolean italic) {
        if (ea == null || ea.isBlank()) return spec(latin, bold, italic);
        if (latin == null || latin.isBlank()) return spec(ea, bold, italic);
        String key = latin + "|" + ea + "|" + bold + "|" + italic;
        Spec s = specs.get(key);
        if (s != null) return s;
        Spec base = spec(latin, bold, italic);
        Spec eas = spec(ea, bold, italic);
        Face eaFace = eas.face.hasGlyph(0x3042) ? eas.face : eas.eaFace;
        // 欧文フォント自体が和文でも、主の書体はそのまま
        Face face = base.face;
        float asc = Math.max(base.ascent, eas.ascent), desc = Math.max(base.descent, eas.descent);
        s = new Spec(face, eaFace, bold, italic, asc, desc, latin);
        s.textAscent = base.textAscent;
        s.textDescent = base.textDescent;
        specs.put(key, s);
        return s;
    }

    /** CJK の文字 (漢字、かな、全角形、CJK の記号など) か。 */
    public static boolean isCjk(int cp) {
        return (cp >= 0x2E80 && cp <= 0x9FFF) || (cp >= 0xF900 && cp <= 0xFAFF) || (cp >= 0xFF00 && cp <= 0xFFEF)
                || (cp >= 0x3000 && cp <= 0x303F) || (cp >= 0x20000 && cp <= 0x2FFFF);
    }

    /** 96 dpi での最大数字幅 (px)。Excel が列幅の単位に使う。 */
    public int maxDigitWidthPx(String name, float sizePt) {
        int measured = PrinterModel.measuredMdw(name, sizePt, 96);
        if (measured > 0) return measured;
        String n = norm(name);
        int size = Math.round(sizePt);
        // Excel で実測した値 (GDI のヒンティング後の幅)
        int[] tbl = lookupLoose(KNOWN_MDW, n);
        if (tbl != null && size >= 6 && size < 6 + tbl.length) return tbl[size - 6];
        float ppem = Math.round(sizePt * 96f / 72f);
        double[] g = gdiMetrics(name);
        if (g != null) return (int) Math.max(1, Math.round(g[2] * ppem));
        Spec s = spec(name, false, false);
        float max = 0;
        for (char c = '0'; c <= '9'; c++) max = Math.max(max, s.face.advance(c));
        return Math.max(1, Math.round(max * ppem));
    }

    /** 実測した 96 dpi の最大数字幅 (サイズ 6〜20pt)。 */
    private static final Map<String, int[]> KNOWN_MDW = new LinkedHashMap<>();

    static {
        //                                    6  7  8  9 10 11 12 13 14 15 16 17 18 19 20
        KNOWN_MDW.put("calibri", new int[]   {4, 5, 6, 6, 7, 7, 8, 9, 10, 10, 11, 12, 13, 13, 14});
        KNOWN_MDW.put("arial", new int[]     {4, 5, 5, 6, 7, 8, 9, 9, 10, 11, 12, 12, 13, 14, 15});
        int[] msp = {4, 5, 6, 6, 7, 8, 8, 9, 10, 10, 11, 12, 13, 13, 14};
        for (String k : new String[] {"ms pゴシック", "ms ゴシック", "ms p明朝", "ms 明朝", "ms pgothic", "ms gothic",
                "ms pmincho", "ms mincho"}) KNOWN_MDW.put(k, msp);
        int[] yu = {4, 5, 6, 7, 7, 8, 9, 10, 10, 11, 12, 13, 13, 14, 15};
        for (String k : new String[] {"游ゴシック", "yu gothic", "游明朝", "yu mincho", "游ゴシック light"}) KNOWN_MDW.put(k, yu);
        int[] mei = {4, 5, 6, 7, 7, 8, 9, 10, 10, 11, 12, 13, 13, 14, 15};
        for (String k : new String[] {"メイリオ", "meiryo", "meiryo ui"}) KNOWN_MDW.put(k, mei);
    }
}
