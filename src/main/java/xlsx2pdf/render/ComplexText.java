package xlsx2pdf.render;

import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.awt.font.GlyphVector;
import java.awt.geom.AffineTransform;
import java.awt.geom.Path2D;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 複雑な文字体系 (デーヴァナーガリーなどのインド系、タイ、アラビア、ヘブライなど) の整形。
 * <p>
 * 本ツールの文字の配置は「1 コードポイント = 1 つの送り幅」で作られている。結合文字・合字・前置の母音記号は
 * この形に合わないので、文字列を「まとまり」(インド系は音節、右から左の文字は語句) に分け、まとまり 1 つを
 * 私用面 (U+100000 以降) の 1 コードポイントに置き換える ({@link #cluster})。まとまりの送り幅と輪郭は、
 * 書体ごとに AWT の整形 (HarfBuzz。GSUB・GPOS を使う) で求める ({@link #shape})。描画は輪郭の塗りで行う
 * (PDF の文字としては残らない)。{@code -Dx2p.complex=0} で無効。
 */
public final class ComplexText {
    private ComplexText() {}

    /** 有効か。 */
    static final boolean ON = !"0".equals(System.getProperty("x2p.complex"));

    /** まとまりに割り当てるコードポイントの始まりと数 (第 16 面の私用領域)。 */
    static final int BASE = 0x100000, MAX = 0xFFFD;

    private static final Map<String, Integer> IDS = new ConcurrentHashMap<>();
    private static final List<String> TEXTS = new ArrayList<>();

    /** まとまりを表すコードポイントか。 */
    public static boolean isCluster(int cp) { return cp >= BASE && cp < BASE + MAX; }

    /** まとまりの元の文字列。 */
    public static String text(int cp) {
        synchronized (TEXTS) {
            int i = cp - BASE;
            return i >= 0 && i < TEXTS.size() ? TEXTS.get(i) : "";
        }
    }

    /** まとまりの最初のコードポイント (書体を選ぶのに使う)。 */
    public static int firstCp(int cp) {
        String t = text(cp);
        for (int i = 0; i < t.length(); ) {
            int c = t.codePointAt(i);
            i += Character.charCount(c);
            if (!Character.isWhitespace(c) && !isMark(c)) return c;
        }
        return t.isEmpty() ? ' ' : t.codePointAt(0);
    }

    /** 右から左に書く文字 (ヘブライ・アラビア・シリアなど)。 */
    static boolean isRtl(int cp) {
        return (cp >= 0x0590 && cp <= 0x08FF) || (cp >= 0xFB1D && cp <= 0xFDFF) || (cp >= 0xFE70 && cp <= 0xFEFF);
    }

    /** 整形の要る文字体系の文字か。 */
    static boolean isComplex(int cp) {
        if (cp < 0x0590) return false;
        return isRtl(cp) || (cp >= 0x0900 && cp <= 0x109F) || (cp >= 0x1780 && cp <= 0x17FF) || (cp >= 0x1A00 && cp <= 0x1AAF)
                || (cp >= 0xA800 && cp <= 0xA8FF) || (cp >= 0xA980 && cp <= 0xAA7F);
    }

    static boolean isMark(int cp) {
        int t = Character.getType(cp);
        return t == Character.NON_SPACING_MARK || t == Character.COMBINING_SPACING_MARK || t == Character.ENCLOSING_MARK
                || cp == 0x200C || cp == 0x200D;
    }

    /** 子音を結合させる記号 (ヴィラーマ・ミャンマーの積み重ね・クメールのコエン)。次の子音も同じまとまりに入る。 */
    private static boolean isJoiner(int cp) {
        switch (cp) {
            case 0x094D: case 0x09CD: case 0x0A4D: case 0x0ACD: case 0x0B4D: case 0x0BCD: case 0x0C4D: case 0x0CCD:
            case 0x0D4D: case 0x0DCA: case 0x1039: case 0x17D2: case 0xA8C4: case 0xA9C0:
                return true;
            default:
                return false;
        }
    }

    private static boolean hasComplex(String s) {
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c >= 0x0590 && isComplex(c)) return true;
        }
        return false;
    }

    private static int id(String t) {
        Integer v = IDS.get(t);
        if (v != null) return v;
        synchronized (TEXTS) {
            v = IDS.get(t);
            if (v != null) return v;
            if (TEXTS.size() >= MAX) return -1;
            TEXTS.add(t);
            v = BASE + TEXTS.size() - 1;
            IDS.put(t, v);
            return v;
        }
    }

    /** 複雑な文字体系の部分を、まとまりごとに 1 つのコードポイントへ置き換える。なければ元の文字列を返す。 */
    public static String cluster(String s) {
        if (!ON || s == null || !hasComplex(s)) return s;
        StringBuilder sb = new StringBuilder(s.length());
        int n = s.length(), i = 0;
        while (i < n) {
            int cp = s.codePointAt(i);
            if (!isComplex(cp)) {
                sb.appendCodePoint(cp);
                i += Character.charCount(cp);
                continue;
            }
            int j = i + Character.charCount(cp);
            if (isRtl(cp)) {
                // 語句: 右から左の文字と、その間の空白・記号。末尾の空白・記号は含めない
                int end = j;
                while (j < n) {
                    int c = s.codePointAt(j);
                    if (c == '\n' || c == '\r' || c == '\t') break;
                    boolean rtl = isRtl(c) || isMark(c);
                    if (!rtl && (Character.isLetterOrDigit(c) || c > 0x2000)) break;
                    j += Character.charCount(c);
                    if (rtl) end = j;
                }
                j = end;
            } else {
                int prev = cp;
                while (j < n) {
                    int c = s.codePointAt(j);
                    if (isMark(c) || (isJoiner(prev) && isComplex(c) && !isRtl(c))) {
                        prev = c;
                        j += Character.charCount(c);
                    } else break;
                }
            }
            int id = id(s.substring(i, j));
            if (id < 0) sb.append(s, i, j);
            else sb.appendCodePoint(id);
            i = j;
        }
        return sb.toString();
    }

    /** まとまりを元の文字列に戻す (文字として取り出すとき)。 */
    public static String plain(String s) {
        if (s == null) return null;
        boolean any = false;
        for (int i = 0; i < s.length() && !any; i++) any = Character.isHighSurrogate(s.charAt(i));
        if (!any) return s;
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < s.length(); ) {
            int cp = s.codePointAt(i);
            i += Character.charCount(cp);
            if (isCluster(cp)) sb.append(text(cp));
            else sb.appendCodePoint(cp);
        }
        return sb.toString();
    }

    /** 整形した結果: 送り幅 (em) と輪郭 (em、基線の原点、y は下向き)。 */
    public static final class Shaped {
        public final float adv;
        public final Path2D path;

        Shaped(float adv, Path2D path) {
            this.adv = adv;
            this.path = path;
        }
    }

    private static final FontRenderContext FRC = new FontRenderContext(null, true, true);
    private static final float EM = 1000f;

    /** まとまり cp を書体 f で整形する。AWT の書体が作れないときは、整形せずに字形を並べる。 */
    static Shaped shape(FontManager.Face f, int cp) {
        String t = text(cp);
        Font awt = f.awtFont();
        if (awt != null) {
            try {
                char[] c = t.toCharArray();
                boolean rtl = isRtl(t.codePointAt(0));
                GlyphVector gv = awt.deriveFont(EM).layoutGlyphVector(FRC, c, 0, c.length, rtl ? Font.LAYOUT_RIGHT_TO_LEFT : Font.LAYOUT_LEFT_TO_RIGHT);
                double adv = gv.getGlyphPosition(gv.getNumGlyphs()).getX();
                Path2D p = new Path2D.Double(gv.getOutline(), AffineTransform.getScaleInstance(1 / EM, 1 / EM));
                return new Shaped((float) (adv / EM), p);
            } catch (RuntimeException e) {
                // 下の簡易な並べ方にする
            }
        }
        Path2D p = new Path2D.Double();
        double x = 0;
        for (int i = 0; i < t.length(); ) {
            int c = t.codePointAt(i);
            i += Character.charCount(c);
            try {
                int gid = f.ttf.getUnicodeCmapLookup().getGlyphId(c);
                var gd = f.ttf.getGlyph() == null ? null : f.ttf.getGlyph().getGlyph(gid);
                double k = 1.0 / f.unitsPerEm;
                if (gd != null) p.append(new Path2D.Double(gd.getPath(), new AffineTransform(k, 0, 0, -k, x, 0)), false);
                x += f.ttf.getAdvanceWidth(gid) * k;
            } catch (java.io.IOException | RuntimeException e) {
                x += 0.5;
            }
        }
        return new Shaped((float) x, p);
    }
}
