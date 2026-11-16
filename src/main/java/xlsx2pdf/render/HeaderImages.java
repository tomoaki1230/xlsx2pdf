package xlsx2pdf.render;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/** シートの旧形式 VML 描画から、ヘッダー/フッターの図 (&amp;G) を読み込む。 */
final class HeaderImages {
    private HeaderImages() {}

    /** VML の style 属性の幅・高さ (単位付き)。 */
    private static final Pattern SIZE = Pattern.compile("(width|height)\\s*:\\s*([0-9.]+)\\s*(pt|in|px|mm|cm)?");

    /**
     * ヘッダー/フッターの図。
     *
     * @param image 埋め込んだ画像
     * @param w     幅 (pt)
     * @param h     高さ (pt)
     */
    record Img(PDImageXObject image, double w, double h) {}

    /**
     * シートのヘッダー/フッター用の図を読み込む。
     *
     * @return VML の図形 ID (LH, CF など) から図への対応。読めない図は含めない
     */
    static Map<String, Img> load(RenderContext ctx, XSSFSheet sheet) {
        Map<String, Img> out = new HashMap<>();
        try {
            if (!sheet.getCTWorksheet().isSetLegacyDrawingHF()) return out;
            String rid = sheet.getCTWorksheet().getLegacyDrawingHF().getId();
            PackagePart sp = sheet.getPackagePart();
            PackageRelationship rel = sp.getRelationship(rid);
            if (rel == null) return out;
            PackagePart vml = sp.getRelatedPart(rel);
            Element root;
            try (InputStream in = vml.getInputStream()) {
                root = parseVml(in.readAllBytes());
            }
            NodeList shapes = root.getElementsByTagNameNS("*", "shape");
            for (int i = 0; i < shapes.getLength(); i++) {
                Element s = (Element) shapes.item(i);
                String id = s.getAttribute("id");
                double[] wh = styleSize(s.getAttribute("style"));
                Element im = null;
                for (Node n = s.getFirstChild(); n != null; n = n.getNextSibling()) {
                    if (n instanceof Element e && "imagedata".equals(e.getLocalName())) im = e;
                }
                if (im == null) continue;
                String r = im.getAttributeNS("urn:schemas-microsoft-com:office:office", "relid");
                if (r.isEmpty()) r = im.getAttribute("o:relid");
                PackageRelationship ir = vml.getRelationship(r);
                if (ir == null) continue;
                PackagePart ip = vml.getRelatedPart(ir);
                byte[] data;
                try (InputStream in = ip.getInputStream()) {
                    data = in.readAllBytes();
                }
                // 切り抜き・グレースケール・明るさ (VML の imagedata の属性。exp122 img_crop・img_color)
                String key = ip.getPartName().getName();
                byte[] adj = adjust(data, im);
                if (adj != data) key = key + "#" + im.getAttribute("cropleft") + im.getAttribute("croptop") + im.getAttribute("cropright")
                        + im.getAttribute("cropbottom") + im.getAttribute("grayscale") + im.getAttribute("gain") + im.getAttribute("blacklevel");
                PDImageXObject img = Images.toPdf(ctx, key, adj, ip.getPartName().getName(), wh[0], wh[1]);
                if (img != null) out.put(id, new Img(img, wh[0], wh[1]));
            }
        } catch (Exception e) {
            // 読めないヘッダーの図は無視する
            Warnings.add("ヘッダー・フッターの画像を読めませんでした", e);
        }
        return out;
    }

    /** VML の分数 ("13107f" = 13107/65536、または小数)。なければ dflt。 */
    static double frac(String v, double dflt) {
        if (v == null || v.isBlank()) return dflt;
        v = v.strip();
        try {
            return v.endsWith("f") ? Double.parseDouble(v.substring(0, v.length() - 1)) / 65536.0 : Double.parseDouble(v);
        } catch (NumberFormatException e) {
            return dflt;
        }
    }

    /**
     * 切り抜き (cropleft/top/right/bottom)・グレースケール (grayscale="t"、BT.709 の輝度)・明るさとコントラスト (gain・blacklevel)
     * を画像に当てる。何もなければ元のデータを返す。明るさ・コントラストは Excel の印刷の実測から「gain × 元の値 + 切片」とし、
     * 切片は 2 点の実測 (gain 0.3・blacklevel 0.35 → 0.804、gain 0.6・blacklevel 0.3 → 0.678) から求めた近似。
     */
    static byte[] adjust(byte[] data, Element im) {
        double cl = frac(im.getAttribute("cropleft"), 0), ct = frac(im.getAttribute("croptop"), 0);
        double cr = frac(im.getAttribute("cropright"), 0), cb = frac(im.getAttribute("cropbottom"), 0);
        boolean gray = "t".equals(im.getAttribute("grayscale")) || "true".equals(im.getAttribute("grayscale"));
        boolean bc = !im.getAttribute("gain").isEmpty() || !im.getAttribute("blacklevel").isEmpty();
        if (cl <= 0 && ct <= 0 && cr <= 0 && cb <= 0 && !gray && !bc) return data;
        try {
            java.awt.image.BufferedImage src = javax.imageio.ImageIO.read(new java.io.ByteArrayInputStream(data));
            if (src == null) return data;
            int w = src.getWidth(), h = src.getHeight();
            int x0 = (int) Math.round(w * Math.max(0, cl)), y0 = (int) Math.round(h * Math.max(0, ct));
            int x1 = w - (int) Math.round(w * Math.max(0, cr)), y1 = h - (int) Math.round(h * Math.max(0, cb));
            if (x1 - x0 < 1 || y1 - y0 < 1) return data;
            double gain = frac(im.getAttribute("gain"), 1), black = frac(im.getAttribute("blacklevel"), 0);
            double icpt = bc ? 255 * (2.186 * black + 0.0557 * (1 - gain)) : 0;
            java.awt.image.BufferedImage out = new java.awt.image.BufferedImage(x1 - x0, y1 - y0, java.awt.image.BufferedImage.TYPE_INT_ARGB);
            for (int y = y0; y < y1; y++) {
                for (int x = x0; x < x1; x++) {
                    int argb = src.getRGB(x, y);
                    int a = argb >>> 24, r = (argb >> 16) & 255, g = (argb >> 8) & 255, b = argb & 255;
                    if (gray) {
                        int l = (int) Math.round(0.2126 * r + 0.7152 * g + 0.0722 * b);
                        r = g = b = l;
                    }
                    if (bc) {
                        r = clamp(gain * r + icpt); g = clamp(gain * g + icpt); b = clamp(gain * b + icpt);
                    }
                    out.setRGB(x - x0, y - y0, (a << 24) | (r << 16) | (g << 8) | b);
                }
            }
            java.io.ByteArrayOutputStream bo = new java.io.ByteArrayOutputStream();
            javax.imageio.ImageIO.write(out, "png", bo);
            return bo.toByteArray();
        } catch (Exception e) {
            return data;
        }
    }

    private static int clamp(double v) { return (int) Math.max(0, Math.min(255, Math.round(v))); }

    /** VML は整形式の XML でないことがある (&lt;br&gt; など) ため、よくある崩れを直してから解析する。 */
    static Element parseVml(byte[] data) throws Exception {
        String s = new String(data, StandardCharsets.UTF_8);
        s = s.replaceAll("<br>", "<br/>");
        return Xml.parse(new java.io.ByteArrayInputStream(s.getBytes(StandardCharsets.UTF_8)));
    }

    /**
     * VML の style 属性から幅と高さを読む。
     *
     * @return {幅, 高さ} (pt)。単位がなければ px とみなす
     */
    static double[] styleSize(String style) {
        double w = 0, h = 0;
        Matcher m = SIZE.matcher(style);
        while (m.find()) {
            double v = Double.parseDouble(m.group(2));
            String u = m.group(3) == null ? "px" : m.group(3);
            v = switch (u) {
                case "in" -> v * 72;
                case "px" -> v * 0.75;
                case "mm" -> v * 72 / 25.4;
                case "cm" -> v * 72 / 2.54;
                default -> v;
            };
            if (m.group(1).equals("width")) w = v; else h = v;
        }
        return new double[] {w, h};
    }
}
