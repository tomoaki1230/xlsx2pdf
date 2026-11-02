package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.awt.geom.RoundRectangle2D;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/** 旧形式 (VML) のフォームコントロール (チェックボックス、オプションボタン、ラベル、グループボックス、ボタンなど)。 */
final class VmlControls {
    /**
     * VML の色 ("#rrggbb"、"window [65]" のようなシステム色、色の名前) を読む。読めなければ白
     * (塗りの既定。window はウィンドウの背景色 = 白)。
     */
    static Color vmlColor(String v) {
        if (v == null) return Color.WHITE;
        String t = v.trim();
        int sp = t.indexOf(' ');
        if (sp > 0) t = t.substring(0, sp);
        if (t.startsWith("#") && (t.length() == 7 || t.length() == 4)) {
            try {
                if (t.length() == 4) t = "#" + t.charAt(1) + t.charAt(1) + t.charAt(2) + t.charAt(2) + t.charAt(3) + t.charAt(3);
                return new Color(Integer.parseInt(t.substring(1), 16));
            } catch (NumberFormatException e) {
                return Color.WHITE;
            }
        }
        return switch (t.toLowerCase(java.util.Locale.ROOT)) {
            case "black", "windowtext" -> Color.BLACK;
            case "silver", "buttonface" -> new Color(0xC0C0C0);
            case "gray" -> Color.GRAY;
            default -> Color.WHITE;
        };
    }

    private VmlControls() {}

    /** 1 つのフォームコントロール。 */
    static final class Control {
        /** 種類 (ClientData の ObjectType: Checkbox, Radio, Label, GBox, Button, Drop, List など)。 */
        String type;
        /** アンカーの開始・終了のセル (0 始まりの列・行)。 */
        int c0, r0, c1, r1;
        /** アンカーのセル内オフセット (96 dpi の画面ピクセル)。 */
        double c0Off, r0Off, c1Off, r1Off;
        /** チェックされているか。 */
        boolean checked;
        /** 印刷するか。 */
        boolean print = true;
        /** 表示する文字列。 */
        String text = "";
        /** ActiveX コントロール (Pict) の印刷用の画像 (v:imagedata。なければ null) とパーツ名。 */
        byte[] image;
        String imageName;
        /** Pict の図形の塗り (filled="t" のとき。なければ null)。OLE オブジェクトは白で塗り、下のセルの文字を隠す。 */
        Color fill;
        /** コメントの文字の大きさ (pt)。 */
        float noteSize = 9;
        /** 画像を PDF にするための描画環境。 */
        RenderContext ctx;
        /** 立体表示なし。 */
        boolean noThreeD;
        /** 文字のフォント。 */
        FontManager.Spec fontSpec;

        /** アンカーのずれを、セルの大きさ size までに切り詰める。 */
        private static double off(double d, double size) {
            return Math.min(d, Math.max(0, size));
        }

        /** コントロールを描画する (アンカーの位置と大きさに合わせる)。 */
        void paint(PdfCanvas cv, DrawingPainter dp, SheetLayout L) throws IOException {
            // アンカーのずれ (画面 px) はセルの大きさまで (Excel はセルを超えるずれを切り詰める。
            // チェックボックスは高さ 16px の行に 23px のずれ)
            double x0 = L.x(c0) + off(c0Off * 0.75 * L.xScale, L.cw(c0)), y0 = L.y(r0) + off(r0Off * 0.75 * L.yScale, L.rh(r0));
            double x1 = L.x(c1) + off(c1Off * 0.75 * L.xScale, L.cw(c1)), y1 = L.y(r1) + off(r1Off * 0.75 * L.yScale, L.rh(r1));
            double w = x1 - x0, h = y1 - y0;
            if (w <= 0 || h <= 0) return;
            FontManager.Spec sp = fontSpec;
            float size = 9f;
            switch (type) {
                case "Note" -> {
                    // コメント: 薄い黄色の背景、黒い枠、左上から折り返した文字
                    cv.fillRect(x0, y0, w, h, new Color(255, 255, 225));
                    cv.stroke(new Rectangle2D.Double(x0, y0, w, h), new PdfCanvas.Stroke(0.75f), Color.BLACK);
                    if (sp == null || text.isEmpty()) return;
                    float fs = noteSize;
                    double pad = 2, lineH = (sp.ascent + sp.descent) * fs * 1.1, y = y0 + pad + sp.ascent * fs;
                    cv.save();
                    cv.clipRect(x0, y0, w, h);
                    for (String para : text.split("\\r?\\n")) {
                        StringBuilder line = new StringBuilder();
                        for (int k = 0; k < para.length(); ) {
                            int cp = para.codePointAt(k);
                            String ch = new String(Character.toChars(cp));
                            if (line.length() > 0 && sp.width(line + ch, fs, cv.fonts) > w - 2 * pad) {
                                cv.text(line.toString(), x0 + pad, y, sp, fs, Color.BLACK, 1, 0);
                                y += lineH;
                                line.setLength(0);
                            }
                            line.append(ch);
                            k += Character.charCount(cp);
                        }
                        cv.text(line.toString(), x0 + pad, y, sp, fs, Color.BLACK, 1, 0);
                        y += lineH;
                    }
                    cv.restore();
                }
                case "Pict" -> {
                    // ActiveX コントロール: Excel は保存された見た目の画像 (EMF など) を印刷する
                    if (image == null || ctx == null) return;
                    org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject img =
                            Images.toPdf(ctx, "vml:" + imageName + ":" + Math.round(w) + "x" + Math.round(h), image, imageName, w, h);
                    if (img != null) {
                        if (fill != null) cv.fillRect(x0, y0, w, h, fill);
                        cv.image(img, x0, y0, w, h, 1f);
                    }
                }
                case "Checkbox", "Radio" -> {
                    // 箱の大きさ 9pt
                    double box = 9.0;
                    double bx = x0 + 1.5, by = y0 + (h - box) / 2;
                    PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.75f);
                    if (type.equals("Checkbox")) {
                        cv.fillRect(bx, by, box, box, Color.WHITE);
                        cv.stroke(new Rectangle2D.Double(bx, by, box, box), st, Color.BLACK);
                        if (checked) {
                            Path2D p = new Path2D.Double();
                            p.moveTo(bx + box * 0.22, by + box * 0.5);
                            p.lineTo(bx + box * 0.42, by + box * 0.76);
                            p.lineTo(bx + box * 0.8, by + box * 0.22);
                            PdfCanvas.Stroke ck = new PdfCanvas.Stroke(1.3f);
                            cv.stroke(p, ck, Color.BLACK);
                        }
                    } else {
                        Ellipse2D e = new Ellipse2D.Double(bx, by, box, box);
                        cv.fill(e, Color.WHITE);
                        cv.stroke(e, st, Color.BLACK);
                        if (checked) cv.fill(new Ellipse2D.Double(bx + box * 0.28, by + box * 0.28, box * 0.44, box * 0.44), Color.BLACK);
                    }
                    if (!text.isEmpty() && sp != null) {
                        double base = y0 + h / 2 + (sp.ascent - sp.descent) * size / 2;
                        cv.save();
                        cv.clipRect(x0, y0, w, h);
                        cv.text(text, bx + 12, base, sp, size, Color.BLACK, 1, 0); // 文字の位置は箱の大きさによらない
                        cv.restore();
                    }
                }
                case "Label" -> {
                    if (!text.isEmpty() && sp != null) {
                        cv.save();
                        cv.clipRect(x0, y0, w, h);
                        cv.text(text, x0 + 1, y0 + sp.ascent * size + 1, sp, size, Color.BLACK, 1, 0);
                        cv.restore();
                    }
                }
                case "GBox" -> {
                    PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.75f);
                    double ty = y0 + size * 0.6;
                    cv.stroke(new Rectangle2D.Double(x0, ty, w, h - (ty - y0)), st, new Color(0x808080));
                    if (!text.isEmpty() && sp != null) {
                        float tw = sp.width(text, size, cv.fonts);
                        cv.fillRect(x0 + 6, y0, tw + 4, size * 1.2, Color.WHITE);
                        cv.text(text, x0 + 8, y0 + sp.ascent * size, sp, size, Color.BLACK, 1, 0);
                    }
                }
                case "Button" -> {
                    RoundRectangle2D rr = new RoundRectangle2D.Double(x0, y0, w, h, 2, 2);
                    cv.fill(rr, new Color(0xF0F0F0));
                    cv.stroke(rr, new PdfCanvas.Stroke(0.75f), new Color(0x707070));
                    if (!text.isEmpty() && sp != null) {
                        float tw = sp.width(text, size, cv.fonts);
                        cv.text(text, x0 + (w - tw) / 2, y0 + h / 2 + (sp.ascent - sp.descent) * size / 2, sp, size, Color.BLACK, 1, 0);
                    }
                }
                case "Drop", "List" -> {
                    cv.fillRect(x0, y0, w, h, Color.WHITE);
                    cv.stroke(new Rectangle2D.Double(x0, y0, w, h), new PdfCanvas.Stroke(0.75f), new Color(0x808080));
                }
                default -> { }
            }
        }
    }

    /**
     * シートの旧形式描画からフォームコントロールを読み込む。コメント (Note) は既定で印刷しないため除く。
     *
     * @return コントロールの一覧。読めなければ空
     */
    static List<Control> load(RenderContext ctx, XSSFSheet sheet) {
        List<Control> out = new ArrayList<>();
        try {
            if (!sheet.getCTWorksheet().isSetLegacyDrawing()) return out;
            String rid = sheet.getCTWorksheet().getLegacyDrawing().getId();
            PackagePart sp = sheet.getPackagePart();
            PackageRelationship rel = sp.getRelationship(rid);
            if (rel == null) return out;
            PackagePart vml = sp.getRelatedPart(rel);
            Element root;
            try (InputStream in = vml.getInputStream()) {
                root = HeaderImages.parseVml(in.readAllBytes());
            }
            FontManager.Spec font = ctx.fonts.spec(ctx.japanese ? "ＭＳ Ｐゴシック" : "Tahoma", false, false);
            boolean asDisplayed = sheet.getCTWorksheet().isSetPageSetup() && sheet.getCTWorksheet().getPageSetup().isSetCellComments()
                    && sheet.getCTWorksheet().getPageSetup().getCellComments() == org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellComments.AS_DISPLAYED;
            NodeList shapes = root.getElementsByTagNameNS("*", "shape");
            for (int i = 0; i < shapes.getLength(); i++) {
                Element s = (Element) shapes.item(i);
                Element cd = null;
                for (Node n = s.getFirstChild(); n != null; n = n.getNextSibling())
                    if (n instanceof Element e && "ClientData".equals(e.getLocalName())) cd = e;
                if (cd == null) continue;
                String type = cd.getAttribute("ObjectType");
                // コメントは既定で印刷しない。「画面表示イメージ」(cellComments="asDisplayed") で表示中のものだけ印刷する
                if (type.equals("Note") && !(asDisplayed && Xml.child(cd, "Visible") != null)) continue;
                Control c = new Control();
                c.type = type;
                c.fontSpec = font;
                String style = s.getAttribute("style");
                if (style.contains("visibility:hidden")) c.print = false;
                for (Element e : Xml.children(cd)) {
                    String n = e.getLocalName();
                    String v = e.getTextContent().trim();
                    switch (n) {
                        case "Anchor" -> {
                            String[] p = v.split(",");
                            if (p.length == 8) {
                                c.c0 = Integer.parseInt(p[0].trim()); c.c0Off = Double.parseDouble(p[1].trim());
                                c.r0 = Integer.parseInt(p[2].trim()); c.r0Off = Double.parseDouble(p[3].trim());
                                c.c1 = Integer.parseInt(p[4].trim()); c.c1Off = Double.parseDouble(p[5].trim());
                                c.r1 = Integer.parseInt(p[6].trim()); c.r1Off = Double.parseDouble(p[7].trim());
                            }
                        }
                        case "Checked" -> c.checked = !v.equals("0");
                        case "PrintObject" -> c.print = !v.equalsIgnoreCase("False");
                        case "NoThreeD" -> c.noThreeD = true;
                        default -> { }
                    }
                }
                if (type.equals("Pict")) {
                    // 塗り
                    if ("t".equals(s.getAttribute("filled")) || "true".equals(s.getAttribute("filled"))) c.fill = vmlColor(s.getAttribute("fillcolor"));
                    // v:fill の opacity="0" や on="f" は塗らない
                    NodeList vf = s.getElementsByTagNameNS("urn:schemas-microsoft-com:vml", "fill");
                    if (c.fill != null && vf.getLength() > 0) {
                        Element fe = (Element) vf.item(0);
                        String op = fe.getAttribute("opacity").trim(), on = fe.getAttribute("on");
                        double o = 1;
                        try {
                            if (op.endsWith("f")) o = Double.parseDouble(op.substring(0, op.length() - 1)) / 65536.0;
                            else if (!op.isEmpty()) o = Double.parseDouble(op);
                        } catch (NumberFormatException ignored) {
                            o = 1;
                        }
                        if (o <= 0.001 || "f".equals(on) || "false".equals(on)) c.fill = null;
                    }
                    NodeList im = s.getElementsByTagNameNS("*", "imagedata");
                    if (im.getLength() > 0) {
                        Element ie = (Element) im.item(0);
                        String relid = ie.getAttributeNS("urn:schemas-microsoft-com:office:office", "relid");
                        if (relid.isEmpty()) relid = ie.getAttribute("o:relid");
                        PackageRelationship ir = relid.isEmpty() ? null : vml.getRelationship(relid);
                        if (ir != null) {
                            PackagePart ip = vml.getRelatedPart(ir);
                            try (InputStream in = ip.getInputStream()) {
                                c.image = in.readAllBytes();
                            }
                            c.imageName = ip.getPartName().getName();
                            c.ctx = ctx;
                        }
                    }
                }
                NodeList tb = s.getElementsByTagNameNS("*", "textbox");
                if (tb.getLength() > 0) c.text = tb.item(0).getTextContent().replaceAll("\\s+", " ").trim();
                if (type.equals("Note")) {
                    // コメントの文字はコメントのパーツにある (VML の textbox は空のことが多い)
                    Element rw = Xml.child(cd, "Row"), cl = Xml.child(cd, "Column");
                    if (rw != null && cl != null) {
                        var cm = sheet.getCellComment(new org.apache.poi.ss.util.CellAddress(
                                Integer.parseInt(rw.getTextContent().trim()), Integer.parseInt(cl.getTextContent().trim())));
                        if (cm != null && cm.getString() != null) {
                            c.text = cm.getString().getString();
                            var rs = (org.apache.poi.xssf.usermodel.XSSFRichTextString) cm.getString();
                            if (rs.numFormattingRuns() > 0 && rs.getFontOfFormattingRun(0) != null)
                                c.noteSize = rs.getFontOfFormattingRun(0).getFontHeightInPoints();
                        }
                    }
                }
                out.add(c);
            }
        } catch (Exception e) {
            // 読めない描画は無視する
        }
        return out;
    }
}
