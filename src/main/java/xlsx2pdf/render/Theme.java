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
        if (tt == null) return;
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
        // Excel の濃淡は 4 桁の値として効く
        tint = Math.round(tint * 10000) / 10000.0;
        double[] hls = rgbToHls(c);
        if (tint < 0) hls[1] = hls[1] * (1 + tint);
        else hls[1] = hls[1] * (1 - tint) + tint;
        return hlsToRgb(hls, c.getAlpha());
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
                case "tint" -> c = new Color(clamp(1 - (1 - c.getRed() / 255.0) * v), clamp(1 - (1 - c.getGreen() / 255.0) * v),
                        clamp(1 - (1 - c.getBlue() / 255.0) * v));
                case "shade" -> c = new Color(clamp(c.getRed() / 255.0 * v), clamp(c.getGreen() / 255.0 * v),
                        clamp(c.getBlue() / 255.0 * v));
                case "inv" -> c = new Color(255 - c.getRed(), 255 - c.getGreen(), 255 - c.getBlue());
                case "gray" -> { int g = (c.getRed() * 3 + c.getGreen() * 6 + c.getBlue()) / 10; c = new Color(g, g, g); }
                case "comp" -> { double[] h = rgbToHls(c); h[0] = (h[0] + 0.5) % 1; c = hlsToRgb(h, 255); }
                default -> { }
            }
        }
        return new Color(c.getRed(), c.getGreen(), c.getBlue(), alpha);
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
        return switch (v) {
            case "black" -> Color.BLACK;
            case "white" -> Color.WHITE;
            case "red" -> Color.RED;
            case "green" -> new Color(0x008000);
            case "blue" -> Color.BLUE;
            case "yellow" -> Color.YELLOW;
            case "gray", "grey" -> new Color(0x808080);
            case "lightGray", "ltGray" -> new Color(0xD3D3D3);
            case "darkGray", "dkGray" -> new Color(0xA9A9A9);
            case "orange" -> new Color(0xFFA500);
            case "cyan", "aqua" -> Color.CYAN;
            case "magenta", "fuchsia" -> Color.MAGENTA;
            case "navy" -> new Color(0x000080);
            case "purple" -> new Color(0x800080);
            case "silver" -> new Color(0xC0C0C0);
            case "maroon" -> new Color(0x800000);
            case "olive" -> new Color(0x808000);
            case "teal" -> new Color(0x008080);
            case "lime" -> new Color(0x00FF00);
            default -> Color.BLACK;
        };
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
