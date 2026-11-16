package xlsx2pdf.render;

import java.awt.Color;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Element;

/** テーマの色・フォント・線のスタイルと、DrawingML の色の解決。 */
public final class Theme {
    /** 配色 (dk1, lt1, accent1 など → 色)。 */
    public final Map<String, Color> scheme = new HashMap<>();
    /** 見出し・本文の欧文フォント。 */
    public String majorLatin = "Calibri Light", minorLatin = "Calibri";
    /** 見出し・本文の東アジアフォント (未指定は空)。 */
    public String majorEa = "", minorEa = "";
    /** 見出し・本文の日本語 (Jpan) フォント (なければ null)。 */
    public String majorJpan, minorJpan;
    /** テーマの部品がない (Apache POI などで作ったブック)。 */
    public boolean missing;
    /** 線のスタイル (lnStyleLst の要素)。 */
    public final List<Element> lineStyles = new ArrayList<>();
    /** 塗りつぶしのスタイル (fillStyleLst の要素)。 */
    public final List<Element> fillStyles = new ArrayList<>();
    /** 効果のスタイル (effectStyleLst の要素)。 */
    public final List<Element> effectStyles = new ArrayList<>();

    /** ブックのテーマを読み込む (テーマがなければ Office 2013 以降の既定値)。 */
    public Theme(XSSFWorkbook wb) {
        // Office 2013 以降の既定値
        put("dk1", 0x000000); put("lt1", 0xFFFFFF); put("dk2", 0x44546A); put("lt2", 0xE7E6E6);
        put("accent1", 0x4472C4); put("accent2", 0xED7D31); put("accent3", 0xA5A5A5); put("accent4", 0xFFC000);
        put("accent5", 0x5B9BD5); put("accent6", 0x70AD47); put("hlink", 0x0563C1); put("folHlink", 0x954F72);
        ThemesTable tt = wb.getTheme();
        if (tt == null) {
            // テーマのないブック (Apache POI などで作ったもの) は、日本語版の Excel の既定のテーマの日本語の書体になる。
            // テーマのフォント (scheme) のセルは、名前 (Calibri・ＭＳ ゴシックなど) によらず游ゴシックで描かれる (exp108)
            missing = true;
            return;
        }
        try (InputStream in = tt.getPackagePart().getInputStream()) {
            Element root = Xml.parse(in);
            Element te = Xml.child(root, "themeElements");
            Element cs = Xml.child(te, "clrScheme");
            if (cs != null) {
                for (Element e : Xml.children(cs)) {
                    Color c = resolve(Xml.firstChild(e), null);
                    if (c != null) scheme.put(e.getLocalName(), c);
                }
            }
            Element fs = Xml.child(te, "fontScheme");
            if (fs != null) {
                Element maj = Xml.child(fs, "majorFont"), min = Xml.child(fs, "minorFont");
                if (maj != null) {
                    majorLatin = Xml.attr(Xml.child(maj, "latin"), "typeface", majorLatin);
                    majorEa = Xml.attr(Xml.child(maj, "ea"), "typeface", "");
                    majorJpan = scriptFont(maj, "Jpan");
                }
                if (min != null) {
                    minorLatin = Xml.attr(Xml.child(min, "latin"), "typeface", minorLatin);
                    minorEa = Xml.attr(Xml.child(min, "ea"), "typeface", "");
                    minorJpan = scriptFont(min, "Jpan");
                }
            }
            Element fmt = Xml.child(te, "fmtScheme");
            if (fmt != null) {
                Element ln = Xml.child(fmt, "lnStyleLst");
                if (ln != null) lineStyles.addAll(Xml.children(ln));
                Element fl = Xml.child(fmt, "fillStyleLst");
                if (fl != null) fillStyles.addAll(Xml.children(fl));
                Element ef = Xml.child(fmt, "effectStyleLst");
                if (ef != null) effectStyles.addAll(Xml.children(ef));
            }
        } catch (Exception e) {
            // 既定値のまま
        }
    }

    /** フォント定義から、指定の文字体系 (Jpan など) のフォントを探す (なければ null)。 */
    private static String scriptFont(Element font, String script) {
        for (Element f : Xml.children(font, "font")) {
            if (script.equals(f.getAttribute("script"))) return f.getAttribute("typeface");
        }
        return null;
    }

    private void put(String k, int rgb) { scheme.put(k, new Color(rgb)); }

    /** "+mn-lt" や "+mj-ea" のようなテーマフォントの参照を、フォント名にする。 */
    public String fontRef(String typeface, boolean ea) {
        if (typeface == null || typeface.isEmpty()) return null;
        return switch (typeface) {
            case "+mn-lt" -> minorLatin;
            case "+mj-lt" -> majorLatin;
            case "+mn-ea" -> minorEa.isEmpty() ? (minorJpan != null ? minorJpan : minorLatin) : minorEa;
            case "+mj-ea" -> majorEa.isEmpty() ? (majorJpan != null ? majorJpan : majorLatin) : majorEa;
            default -> typeface;
        };
    }

    // ---------------- SpreadsheetML の色 ----------------

    /** セルの書式のテーマ色番号に対応する配色名 (0 = lt1、1 = dk1 の順。DrawingML と逆)。 */
    private static final String[] THEME_IDX = {"lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4",
            "accent5", "accent6", "hlink", "folHlink"};

    /**
     * セル・フォントの色を、濃淡 (tint) を適用して解決する。
     *
     * @return 色。自動・色なしは null
     */
    public Color color(XSSFColor c) {
        if (c == null) return null;
        if (c.isAuto() && !c.isRGB() && !c.isThemed()) return null;
        // システムの前景色・背景色 (POI は黒の RGB を返す) は自動とする
        if (c.isIndexed() && !c.isRGB() && !c.isThemed() && (c.getIndex() == 64 || c.getIndex() == 65)) return null;
        Color base = null;
        if (c.isThemed()) {
            int t = c.getTheme();
            if (t >= 0 && t < THEME_IDX.length) base = scheme.get(THEME_IDX[t]);
        }
        if (base == null) {
            byte[] rgb = c.getRGB();
            if (rgb == null) return null;
            base = new Color(rgb[0] & 255, rgb[1] & 255, rgb[2] & 255);
        }
        double tint = c.getTint();
        if (tint != 0) base = applyTint(base, tint);
        return base;
    }

    /** Excel の濃淡 (tint): HLS の輝度を調整する (負なら暗く、正なら明るく)。 */
    public static Color applyTint(Color c, double tint) {
        // Excel の濃淡は 4 桁の値として効く。HLS は Windows の整数の変換 (HLSMAX = 240) で、輝度は切り捨て
        // (Excel の印刷 exp77: accent1 +0.6 → 83CCEB、accent2 −0.25 → BE5014 と一致。小数の HLS では 1〜3 ずれる)
        tint = Math.round(tint * 10000) / 10000.0;
        int[] hls = winRgbToHls(c.getRed(), c.getGreen(), c.getBlue());
        double l = tint < 0 ? hls[1] * (1 + tint) : hls[1] * (1 - tint) + HLSMAX * tint;
        int[] rgb = winHlsToRgb(hls[0], (int) Math.floor(l), hls[2]);
        return new Color(rgb[0], rgb[1], rgb[2], c.getAlpha());
    }

    /** Windows の HLS の最大値 (色相・輝度・彩度は 0〜240 の整数)。 */
    private static final int HLSMAX = 240, RGBMAX = 255;

    /** Windows の RGB → HLS (整数の計算。KB29240 の方式)。 */
    static int[] winRgbToHls(int r, int g, int b) {
        int mx = Math.max(r, Math.max(g, b)), mn = Math.min(r, Math.min(g, b));
        int l = ((mx + mn) * HLSMAX + RGBMAX) / (2 * RGBMAX);
        if (mx == mn) return new int[] {HLSMAX * 2 / 3, l, 0};
        int s = l <= HLSMAX / 2 ? ((mx - mn) * HLSMAX + (mx + mn) / 2) / (mx + mn)
                : ((mx - mn) * HLSMAX + (2 * RGBMAX - mx - mn) / 2) / (2 * RGBMAX - mx - mn);
        int rd = ((mx - r) * (HLSMAX / 6) + (mx - mn) / 2) / (mx - mn);
        int gd = ((mx - g) * (HLSMAX / 6) + (mx - mn) / 2) / (mx - mn);
        int bd = ((mx - b) * (HLSMAX / 6) + (mx - mn) / 2) / (mx - mn);
        int h = r == mx ? bd - gd : g == mx ? HLSMAX / 3 + rd - bd : 2 * HLSMAX / 3 + gd - rd;
        if (h < 0) h += HLSMAX;
        if (h > HLSMAX) h -= HLSMAX;
        return new int[] {h, l, s};
    }

    private static int hueToRgb(int n1, int n2, int h) {
        if (h < 0) h += HLSMAX;
        if (h > HLSMAX) h -= HLSMAX;
        if (h < HLSMAX / 6) return n1 + ((n2 - n1) * h + HLSMAX / 12) / (HLSMAX / 6);
        if (h < HLSMAX / 2) return n2;
        if (h < HLSMAX * 2 / 3) return n1 + ((n2 - n1) * (HLSMAX * 2 / 3 - h) + HLSMAX / 12) / (HLSMAX / 6);
        return n1;
    }

    /** Windows の HLS → RGB (整数の計算)。 */
    static int[] winHlsToRgb(int h, int l, int s) {
        l = Math.max(0, Math.min(HLSMAX, l));
        if (s == 0) { int v = l * RGBMAX / HLSMAX; return new int[] {v, v, v}; }
        int m2 = l <= HLSMAX / 2 ? (l * (HLSMAX + s) + HLSMAX / 2) / HLSMAX : l + s - (l * s + HLSMAX / 2) / HLSMAX;
        int m1 = 2 * l - m2;
        int r = (hueToRgb(m1, m2, h + HLSMAX / 3) * RGBMAX + HLSMAX / 2) / HLSMAX;
        int g = (hueToRgb(m1, m2, h) * RGBMAX + HLSMAX / 2) / HLSMAX;
        int b = (hueToRgb(m1, m2, h - HLSMAX / 3) * RGBMAX + HLSMAX / 2) / HLSMAX;
        return new int[] {Math.max(0, Math.min(255, r)), Math.max(0, Math.min(255, g)), Math.max(0, Math.min(255, b))};
    }

    /** RGB を HLS {色相, 輝度, 彩度} (各 0〜1) にする。 */
    static double[] rgbToHls(Color c) {
        double r = c.getRed() / 255.0, g = c.getGreen() / 255.0, b = c.getBlue() / 255.0;
        double max = Math.max(r, Math.max(g, b)), min = Math.min(r, Math.min(g, b));
        double l = (max + min) / 2, h = 0, s = 0;
        if (max != min) {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
            else if (max == g) h = (b - r) / d + 2;
            else h = (r - g) / d + 4;
            h /= 6;
        }
        return new double[] {h, l, s};
    }

    /** HLS {色相, 輝度, 彩度} を RGB にする。 */
    static Color hlsToRgb(double[] hls, int alpha) {
        double h = hls[0], l = Math.max(0, Math.min(1, hls[1])), s = Math.max(0, Math.min(1, hls[2]));
        double r, g, b;
        if (s == 0) {
            r = g = b = l;
        } else {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            double p = 2 * l - q;
            r = hue(p, q, h + 1.0 / 3);
            g = hue(p, q, h);
            b = hue(p, q, h - 1.0 / 3);
        }
        return new Color(clamp(r), clamp(g), clamp(b), alpha);
    }

    private static double hue(double p, double q, double t) {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 0.5) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    private static int clamp(double v) { return (int) Math.max(0, Math.min(255, Math.round(v * 255))); }

    // ---------------- DrawingML の色 ----------------

    /**
     * DrawingML の色要素 (srgbClr, schemeClr, sysClr, prstClr, scrgbClr, hslClr) を、修飾 (lumMod など) を含めて解決する。
     *
     * @param phClr スタイル参照のプレースホルダー色 (phClr)
     * @return 色。色要素でなければ null
     */
    public Color resolve(Element e, Color phClr) {
        if (e == null) return null;
        Color c;
        switch (e.getLocalName()) {
            case "srgbClr" -> c = hex(e.getAttribute("val"));
            case "sysClr" -> {
                String last = e.getAttribute("lastClr");
                c = !last.isEmpty() ? hex(last) : "window".equals(e.getAttribute("val")) ? Color.WHITE : Color.BLACK;
            }
            case "schemeClr" -> {
                String v = e.getAttribute("val");
                c = switch (v) {
                    case "phClr" -> phClr != null ? phClr : Color.BLACK;
                    case "bg1" -> scheme.get("lt1");
                    case "tx1" -> scheme.get("dk1");
                    case "bg2" -> scheme.get("lt2");
                    case "tx2" -> scheme.get("dk2");
                    default -> scheme.get(v);
                };
                if (c == null) c = Color.BLACK;
            }
            case "prstClr" -> c = preset(e.getAttribute("val"));
            case "scrgbClr" -> c = new Color(lin(e, "r"), lin(e, "g"), lin(e, "b"));
            case "hslClr" -> c = hlsToRgb(new double[] {Xml.intAttr(e, "hue", 0) / 21600000.0,
                    Xml.intAttr(e, "lum", 0) / 100000.0, Xml.intAttr(e, "sat", 0) / 100000.0}, 255);
            default -> { return null; }
        }
        return modify(c, e);
    }

    /** scrgbClr の線形 RGB 成分 (0〜100000) を sRGB の 0〜255 にする。 */
    private static int lin(Element e, String a) {
        double v = Xml.intAttr(e, a, 0) / 100000.0;
        double s = v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
        return clamp(s);
    }

    /** 色の修飾 (透明度、輝度・彩度の変更、tint、shade、反転、グレー、補色) を順に適用する。 */
    private static Color modify(Color c, Element e) {
        int alpha = 255;
        for (Element m : Xml.children(e)) {
            double v = Xml.intAttr(m, "val", 0) / 100000.0;
            switch (m.getLocalName()) {
                case "alpha" -> alpha = clamp(v);
                case "alphaMod" -> alpha = clamp(alpha / 255.0 * v);
                case "lumMod" -> { double[] h = rgbToHls(c); h[1] *= v; c = hlsToRgb(h, 255); }
                case "lumOff" -> { double[] h = rgbToHls(c); h[1] += v; c = hlsToRgb(h, 255); }
                case "satMod" -> { double[] h = rgbToHls(c); h[2] *= v; c = hlsToRgb(h, 255); }
                // tint・shade は線形の RGB で (Office と同じ。exp47 smartart2 の accent1 tint 40% が Excel の印刷の色に合う)
                case "tint" -> c = new Color(fromLinear(1 - (1 - toLinear(c.getRed())) * v), fromLinear(1 - (1 - toLinear(c.getGreen())) * v),
                        fromLinear(1 - (1 - toLinear(c.getBlue())) * v));
                case "shade" -> c = new Color(fromLinear(toLinear(c.getRed()) * v), fromLinear(toLinear(c.getGreen()) * v),
                        fromLinear(toLinear(c.getBlue()) * v));
                case "inv" -> c = new Color(255 - c.getRed(), 255 - c.getGreen(), 255 - c.getBlue());
                case "gray" -> { int g = (c.getRed() * 3 + c.getGreen() * 6 + c.getBlue()) / 10; c = new Color(g, g, g); }
                case "comp" -> { double[] h = rgbToHls(c); h[0] = (h[0] + 0.5) % 1; c = hlsToRgb(h, 255); }
                case "alphaOff" -> alpha = Math.max(0, Math.min(255, (int) Math.round(alpha + v * 255)));
                // 色相 (角度は 60000 分の 1 度)・彩度・輝度の設定と加算・倍率
                case "hue" -> { double[] h = rgbToHls(c); h[0] = (Xml.longAttr(m, "val", 0) / 21600000.0) % 1; c = hlsToRgb(h, 255); }
                case "hueOff" -> { double[] h = rgbToHls(c); h[0] = ((h[0] + Xml.longAttr(m, "val", 0) / 21600000.0) % 1 + 1) % 1; c = hlsToRgb(h, 255); }
                case "hueMod" -> { double[] h = rgbToHls(c); h[0] = (h[0] * v % 1 + 1) % 1; c = hlsToRgb(h, 255); }
                case "sat" -> { double[] h = rgbToHls(c); h[2] = v; c = hlsToRgb(h, 255); }
                case "satOff" -> { double[] h = rgbToHls(c); h[2] += v; c = hlsToRgb(h, 255); }
                case "lum" -> { double[] h = rgbToHls(c); h[1] = v; c = hlsToRgb(h, 255); }
                // 赤・緑・青の成分 (線形の RGB で計算する)
                case "red", "redMod", "redOff", "green", "greenMod", "greenOff", "blue", "blueMod", "blueOff" -> {
                    double[] rgb = {toLinear(c.getRed()), toLinear(c.getGreen()), toLinear(c.getBlue())};
                    String n = m.getLocalName();
                    int k = n.startsWith("red") ? 0 : n.startsWith("green") ? 1 : 2;
                    if (n.endsWith("Mod")) rgb[k] *= v; else if (n.endsWith("Off")) rgb[k] += v; else rgb[k] = v;
                    c = new Color(fromLinear(rgb[0]), fromLinear(rgb[1]), fromLinear(rgb[2]));
                }
                case "gamma" -> c = new Color(fromLinear(c.getRed() / 255.0), fromLinear(c.getGreen() / 255.0), fromLinear(c.getBlue() / 255.0));
                case "invGamma" -> c = new Color(clamp(toLinear(c.getRed())), clamp(toLinear(c.getGreen())), clamp(toLinear(c.getBlue())));
                default -> { }
            }
        }
        return new Color(c.getRed(), c.getGreen(), c.getBlue(), alpha);
    }

    /** sRGB の成分 (0〜255) を線形の値 (0〜1) にする。 */
    private static double toLinear(int v) {
        double s = v / 255.0;
        return s <= 0.04045 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
    }

    /** 線形の値 (0〜1) を sRGB の成分 (0〜255) にする。 */
    private static int fromLinear(double v) {
        v = Math.max(0, Math.min(1, v));
        return clamp(v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055);
    }

    /** 16 進の RGB (RRGGBB) を色にする (不正なら黒)。 */
    static Color hex(String s) {
        try {
            return new Color(Integer.parseInt(s, 16));
        } catch (RuntimeException e) {
            return Color.BLACK;
        }
    }

    /** プリセット色名 (prstClr) を色にする (不明なら黒)。 */
    private static Color preset(String v) {
        // DrawingML の ST_PresetColorVal (約 140 色。POI の PresetColor の表)
        org.apache.poi.sl.usermodel.PresetColor p = org.apache.poi.sl.usermodel.PresetColor.valueOfOoxmlId(v);
        if (p == null && v.contains("rey")) p = org.apache.poi.sl.usermodel.PresetColor.valueOfOoxmlId(v.replace("rey", "ray")); // grey → gray
        return p != null && p.color != null ? p.color : Color.BLACK;
    }

    /** 塗りつぶし・線の要素の、最初の色の子要素を解決する (なければ null)。 */
    public Color colorChild(Element parent, Color phClr) {
        if (parent == null) return null;
        for (Element c : Xml.children(parent)) {
            Color r = resolve(c, phClr);
            if (r != null) return r;
        }
        return null;
    }
}
