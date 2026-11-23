package xlsx2pdf.render;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;

import org.apache.fontbox.ttf.CmapLookup;
import org.apache.fontbox.ttf.TTFTable;
import org.apache.fontbox.ttf.TrueTypeFont;

/**
 * フォントファイルの数値から、Windows の GDI が返す値と Excel の画面の寸法を計算する。
 *
 * <p>Excel は実測表を持たず、GDI からフォントの値を受け取って計算している。ここでは同じ値をファイルから求める:
 * tmAscent・tmDescent は VDMX の値 (なければ OS/2 の winAscent・winDescent × ppem / em の四捨五入)、
 * tmInternalLeading は tmAscent + tmDescent − ppem、tmExternalLeading は hhea の lineGap から求める。
 * 数字幅は hdmx の値 (なければ送り幅 × ppem / em の四捨五入)。
 */
final class FontCalc {
    private FontCalc() {}

    /** ファイルから読んだ値。 */
    private record Metrics(double upm, double winAsc, double winDesc, double extLeading, Map<Integer, int[]> vdmx, double digit, Map<Integer, Integer> hdmxDigit, double topExtra) {}

    private static final Map<String, Optional<Metrics>> CACHE = new ConcurrentHashMap<>();

    /** 名前のフォントの値 (ファイルになければ空)。 */
    private static Optional<Metrics> metrics(String name) {
        return CACHE.computeIfAbsent(FontManager.norm(name), n -> {
            try {
                return Optional.ofNullable(FontManager.withRegular(name, FontCalc::read));
            } catch (IOException | RuntimeException e) {
                return Optional.empty();
            }
        });
    }

    private static Metrics read(TrueTypeFont t) throws IOException {
        if (t.getOS2Windows() == null || t.getHorizontalHeader() == null) return null;
        double upm = t.getUnitsPerEm();
        double wa = t.getOS2Windows().getWinAscent(), wd = t.getOS2Windows().getWinDescent();
        var hh = t.getHorizontalHeader();
        // GDI の外部行間: hhea の lineGap のうち、win の高さが hhea の高さを超えた分を除いたもの
        double el = Math.max(0, hh.getLineGap() - ((wa + wd) - (hh.getAscender() - hh.getDescender())));
        CmapLookup cmap = t.getUnicodeCmapLookup();
        int[] gids = new int[10];
        double digit = 0;
        for (int i = 0; i < 10; i++) {
            gids[i] = cmap.getGlyphId('0' + i);
            digit = Math.max(digit, t.getAdvanceWidth(gids[i]));
        }
        // 上揃えで Excel が基線を下げる量 (em): hhea のアセント + 行間の半分が winAscent を超える分 (游ゴシック・游明朝 0.145)
        double te = (hh.getAscender() + hh.getLineGap() / 2.0 - wa) / upm;
        return new Metrics(upm, wa, wd, el, vdmx(t), digit, hdmx(t, gids), te > 0.05 ? te : 0);
    }

    /** VDMX の 1:1 (なければ既定) の組: ppem → {yMax, −yMin}。なければ空。 */
    private static Map<Integer, int[]> vdmx(TrueTypeFont t) throws IOException {
        Map<Integer, int[]> out = new HashMap<>();
        TTFTable tb = t.getTableMap().get("VDMX");
        if (tb == null) return out;
        byte[] b = t.getTableBytes(tb);
        int nRatios = u16(b, 4), off = -1;
        for (int i = 0; i < nRatios; i++) {
            int p = 6 + 4 * i, x = b[p + 1] & 255, y0 = b[p + 2] & 255, y1 = b[p + 3] & 255;
            if (x == 0 && y0 == 0 || x == 1 && y0 <= 1 && 1 <= y1) {
                off = u16(b, 6 + 4 * nRatios + 2 * i);
                break;
            }
        }
        if (off < 0) return out;
        int n = u16(b, off);
        for (int i = 0; i < n; i++) {
            int p = off + 4 + 6 * i;
            out.put(u16(b, p), new int[] {(short) u16(b, p + 2), -(short) u16(b, p + 4)});
        }
        return out;
    }

    /** hdmx の数字の最大幅: ppem → px。なければ空。 */
    private static Map<Integer, Integer> hdmx(TrueTypeFont t, int[] gids) throws IOException {
        Map<Integer, Integer> out = new HashMap<>();
        TTFTable tb = t.getTableMap().get("hdmx");
        if (tb == null) return out;
        byte[] b = t.getTableBytes(tb);
        int n = u16(b, 2), size = (b[4] & 255) << 24 | (b[5] & 255) << 16 | (b[6] & 255) << 8 | b[7] & 255;
        for (int i = 0; i < n; i++) {
            int p = 8 + i * size, m = 0;
            for (int g : gids) m = Math.max(m, b[p + 2 + g] & 255);
            out.put(b[p] & 255, m);
        }
        return out;
    }

    private static int u16(byte[] b, int p) {
        return (b[p] & 255) << 8 | b[p + 1] & 255;
    }

    static int ppem(double size, int dpi) {
        return (int) Math.floor(size * dpi / 72 + 0.5);
    }

    /**
     * Excel の画面の標準行 (px)。ファイルになければ 0。
     * 字面の上下に同じ余白をとる: (tmAscent − 内部行間) + 1 + 2 × max(tmDescent, 内部行間 + 外部行間の半分, 2)。
     */
    static int screenRow(String name, double size, int dpi) {
        Metrics m = metrics(name).orElse(null);
        if (m == null) return 0;
        int p = ppem(size, dpi);
        double s = p / m.upm;
        int[] v = m.vdmx.get(p);
        int a = v != null ? v[0] : round(m.winAsc * s), d = v != null ? v[1] : round(m.winDesc * s);
        int il = a + d - p, el = round(m.extLeading * s);
        return a - il + 1 + 2 * Math.max(Math.max(d, il + (el + Math.max(d - il, 0)) / 2), 2);
    }

    /**
     * 標準行を基準線の上と下に分ける: 上 = (tmAscent − 内部行間) + 余白、下 = 余白 + 1。
     * 異なるフォントが同じ行にあると、Excel の行は上の最大 + 下の最大になる。ファイルになければ null。
     */
    static int[] rowParts(String name, double size, int dpi) {
        Metrics m = metrics(name).orElse(null);
        if (m == null) return null;
        int p = ppem(size, dpi);
        double s = p / m.upm;
        int[] v = m.vdmx.get(p);
        int a = v != null ? v[0] : round(m.winAsc * s), d = v != null ? v[1] : round(m.winDesc * s);
        int il = a + d - p, el = round(m.extLeading * s);
        int pad = Math.max(Math.max(d, il + (el + Math.max(d - il, 0)) / 2), 2);
        return new int[] {a - il + pad, pad + 1};
    }

    /** ppem での tmAscent と tmDescent (フォントファイルが見つからなければ null)。 */
    static int[] ascentDescent(String name, int ppem) {
        Metrics m = metrics(name).orElse(null);
        if (m == null || ppem <= 0) return null;
        int[] v = m.vdmx().get(ppem);
        double k = ppem / m.upm();
        return new int[] {v != null ? v[0] : round(m.winAsc() * k), v != null ? v[1] : round(m.winDesc() * k)};
    }

    /**
     * 上揃えのセルで、Excel が基線を winAscent の位置より下げる量 (プリンタの画素)。hhea の行間 (lineGap) が大きく、
     * アセント + 行間の半分が winAscent を超える書体 (游ゴシック・游明朝) だけ。9・11・14pt で Excel と ±0.1 画素で合う。
     */
    static int topExtraPx(String name, int ppem) {
        Metrics m = metrics(name).orElse(null);
        if (m == null || m.topExtra() <= 0 || ppem <= 0) return 0;
        return Math.max(0, round(m.topExtra() * ppem) - 1);
    }

    /** 字面の高さが em と同じフォントか (winAscent + winDescent = em、外部行間なし。ＭＳ 系・BIZ UD・HG 系など)。 */
    static boolean emHeight(String name) {
        Metrics m = metrics(name).orElse(null);
        return m != null && m.winAsc + m.winDesc == m.upm && m.extLeading == 0;
    }

    /** Excel の画面の最大数字幅 (px)。ファイルになければ 0。 */
    static int digitWidth(String name, double size, int dpi) {
        Metrics m = metrics(name).orElse(null);
        if (m == null) return 0;
        int p = ppem(size, dpi);
        Integer h = m.hdmxDigit.get(p);
        return h != null && h > 0 ? h : Math.max(1, round(m.digit * p / m.upm));   // hdmx の 0 は使わない (列幅の計算で 0 で割らない)
    }

    private static int round(double x) {
        return (int) Math.floor(x + 0.5);
    }
}
