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
                PDImageXObject img = Images.toPdf(ctx, ip.getPartName().getName(), data, ip.getPartName().getName(), wh[0], wh[1]);
                if (img != null) out.put(id, new Img(img, wh[0], wh[1]));
            }
        } catch (Exception e) {
            // 読めないヘッダーの図は無視する
        }
        return out;
    }

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
