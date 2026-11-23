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
            case "gray", "grey" -> Color.GRAY;
            case "red" -> Color.RED;
            case "green" -> new Color(0x008000);
            case "lime" -> Color.GREEN;
            case "blue" -> Color.BLUE;
            case "yellow" -> Color.YELLOW;
            case "aqua", "cyan" -> Color.CYAN;
            case "fuchsia", "magenta" -> Color.MAGENTA;
            case "maroon" -> new Color(0x800000);
            case "navy" -> new Color(0x000080);
            case "olive" -> new Color(0x808000);
            case "purple" -> new Color(0x800080);
            case "teal" -> new Color(0x008080);
            case "white", "window" -> Color.WHITE;
            default -> Color.WHITE;
        };
    }

    private VmlControls() {}

    /** ボタンの地、スクロールバーの溝、値の位置の線の色 (Excel の印刷の実測)。 */
    private static final Color EDGE = new Color(0xAAAAAA), CHECK_LINE = new Color(0x404040);
    /** チェックボックス・オプションボタンの箱の大きさ、枠の左から箱・文字までの距離 (pt)。 */
    static final double CHECK_BOX = 9.5, CHECK_LEFT = 3.5, CHECK_TEXT = 16.25;
    /** ラベルの文字の上の余白、グループの見出しと枠の上の線を上げる量 (pt)。 */
    static final double LABEL_TOP = 3, GBOX_UP = 2.5, GBOX_LINE_UP = 3;
    private static final Color BUTTON = new Color(0xF0F0F0), TRACK = new Color(0xCCCCCC), THUMB = new Color(0x777777);
    /** ドロップダウン・リストボックスの文字の大きさ (pt)、左の余白 (pt)、リストの行の間隔 (文字の大きさの倍)。 */
    static final float CTL_FONT = 9.5f;
    static final double CTL_PAD = 3, LIST_PITCH = 11.5 / 9.5;
    /** 縦のスクロールバーの矢印のボタンの長さ (Dx の倍)。 */
    static final double VSCROLL_BUTTON = 1.4;

    private static Path2D tri(double ax, double ay, double bx, double by, double cx, double cy) {
        Path2D p = new Path2D.Double();
        p.moveTo(ax, ay);
        p.lineTo(bx, by);
        p.lineTo(cx, cy);
        p.closePath();
        return p;
    }

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
        /** 枠の色 (なければ描かない)、文字の大きさ (pt) と色。 */
        Color stroke;
        float textSize = 9f;
        Color textColor = Color.BLACK;
        /** チェックボックスの状態 (0 なし、1 チェック、2 混在)。 */
        int checkState;
        static final double NOTE_PAD_X = 1.11, NOTE_PITCH = 1.133;
        /** コメントのセル (行・列。なければ −1)。 */
        int noteRow = -1, noteCol = -1;
        /** コメントの文字の大きさ (pt)。 */
        float noteSize = 9;
        /** 画像を PDF にするための描画環境。 */
        RenderContext ctx;
        /** 立体表示なし。 */
        boolean noThreeD;
        /** 文字のフォント。 */
        FontManager.Spec fontSpec, listFont;
        /** ドロップダウン・リストボックスの項目 (FmlaRange のセルの表示文字列) と、選んだ項目 (1 始まり。0 はなし)。 */
        List<String> items = new ArrayList<>();
        int sel;
        /** スクロールバー・スピンボタンの値と範囲、横向き、スクロールバー・ボタンの幅 (画面 px)。 */
        int val, min, max = 100;
        boolean horiz;
        double dx = 16;

        /** アンカーのずれを、セルの大きさ size までに切り詰める。 */
        private static double off(double d, double size) {
            return Math.min(d, Math.max(0, size));
        }

        /** コントロールの矩形 (シートの pt)。 */
        Rectangle2D rect(SheetLayout L) {
            double x0 = L.x(c0) + off(c0Off * 0.75 * L.xScale, L.cw(c0)), y0 = L.y(r0) + off(r0Off * 0.75 * L.yScale, L.rh(r0));
            double x1 = L.x(c1) + off(c1Off * 0.75 * L.xScale, L.cw(c1)), y1 = L.y(r1) + off(r1Off * 0.75 * L.yScale, L.rh(r1));
            return new Rectangle2D.Double(x0, y0, Math.max(0, x1 - x0), Math.max(0, y1 - y0));
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
            float size = textSize;
            Color tc = textColor;
            // 塗りと枠 (チェックボックス・オプションボタン・ラベル)
            if (type.equals("Checkbox") || type.equals("Radio") || type.equals("Label")) {
                if (fill != null) cv.fillRect(x0, y0, w, h, fill);
                if (stroke != null) cv.stroke(new Rectangle2D.Double(x0, y0, w, h), new PdfCanvas.Stroke(0.75f), stroke);
            }
            switch (type) {
                case "Note" -> {
                    // コメント: 薄い黄色の背景、黒い枠、左上から折り返した文字
                    // セルの右上の角への矢印 (0.14pt の線と、長さ 6pt ほどの矢じり)
                    if (noteRow >= 0 && noteCol >= 0) {
                        double ax = L.x(noteCol + 1), ay = L.y(noteRow);
                        double dx = ax - x0, dy = ay - y0, len = Math.hypot(dx, dy);
                        if (len > 1) {
                            cv.line(x0, y0, ax, ay, new PdfCanvas.Stroke(0.14f), Color.BLACK);
                            double ux = dx / len, uy = dy / len, al = 6.0, aw = 2.6;
                            Path2D head = new Path2D.Double();
                            head.moveTo(ax, ay);
                            head.lineTo(ax - ux * al - uy * aw, ay - uy * al + ux * aw);
                            head.lineTo(ax - ux * al + uy * aw, ay - uy * al - ux * aw);
                            head.closePath();
                            cv.fill(head, Color.BLACK);
                        }
                    }
                    cv.fillRect(x0, y0, w, h, new Color(255, 255, 225));
                    cv.stroke(new Rectangle2D.Double(x0, y0, w, h), new PdfCanvas.Stroke(0.75f), Color.BLACK);
                    if (sp == null || text.isEmpty()) return;
                    float fs = noteSize;
                    // Excel の印刷: 文字は枠の左から 1.11pt、1 行目の基線は枠の上 + アセント、行の間隔は大きさ × 1.133。
                    // 英文は単語の切れ目で折り返す
                    double padX = NOTE_PAD_X * fs / 9, pitch = fs * NOTE_PITCH;
                    TextLayout.Run run = new TextLayout.Run(text, sp, fs, Color.BLACK);
                    List<TextLayout.Line> lines = TextLayout.layout(List.of(run), (float) Math.max(1, w - 2 * padX), true, true, cv.fonts);
                    double y = y0 + sp.ascent * fs;
                    cv.save();
                    cv.clipRect(x0, y0, w, h);
                    for (TextLayout.Line ln : lines) {
                        TextLayout.drawLine(cv, ln, x0 + padX, y, 0, 0);
                        y += pitch;
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
                    // Excel の印刷: 箱は枠の左から 3.5pt、大きさ 9.5pt、線は濃い灰色 0.6pt。文字は枠の左から 16.25pt
                    double box = CHECK_BOX;
                    double bx = x0 + CHECK_LEFT, by = y0 + (h - box) / 2;
                    PdfCanvas.Stroke st = new PdfCanvas.Stroke(0.6f);
                    if (type.equals("Checkbox")) {
                        cv.fillRect(bx, by, box, box, Color.WHITE);
                        cv.stroke(new Rectangle2D.Double(bx, by, box, box), st, CHECK_LINE);
                        if (checked && checkState == 2) {
                            // 混在: 箱の中を灰色で塗る
                            cv.fillRect(bx + 1.5, by + 1.5, box - 3, box - 3, new Color(0x808080));
                        } else if (checked) {
                            Path2D p = new Path2D.Double();
                            p.moveTo(bx + box * 0.22, by + box * 0.5);
                            p.lineTo(bx + box * 0.42, by + box * 0.76);
                            p.lineTo(bx + box * 0.8, by + box * 0.22);
                            PdfCanvas.Stroke ck = new PdfCanvas.Stroke(0.8f);
                            cv.stroke(p, ck, Color.BLACK);
                        }
                    } else {
                        Ellipse2D e = new Ellipse2D.Double(bx, by, box, box);
                        cv.fill(e, Color.WHITE);
                        cv.stroke(e, st, CHECK_LINE);
                        if (checked) cv.fill(new Ellipse2D.Double(bx + box * 0.28, by + box * 0.28, box * 0.44, box * 0.44), Color.BLACK);
                    }
                    if (!text.isEmpty() && sp != null) {
                        double base = y0 + h / 2 + (sp.ascent - sp.descent) * size / 2;
                        cv.save();
                        cv.clipRect(x0, y0, w, h);
                        cv.text(text, x0 + CHECK_TEXT, base, sp, size, tc, 1, 0); // 文字の位置は箱の大きさによらない
                        cv.restore();
                    }
                }
                case "Label" -> {
                    if (!text.isEmpty() && sp != null) {
                        cv.save();
                        cv.clipRect(x0, y0, w, h);
                        cv.text(text, x0 + 1, y0 + sp.ascent * size + LABEL_TOP, sp, size, tc, 1, 0);
                        cv.restore();
                    }
                }
                case "GBox" -> {
                    PdfCanvas.Stroke st = new PdfCanvas.Stroke(1f);
                    // 見出しの中ほどを枠の上の線にする (見出しは枠の上から 2.5pt 上へ、左から 9.2pt)
                    double ty = y0 + size * 0.6 - GBOX_UP - GBOX_LINE_UP;
                    cv.stroke(new Rectangle2D.Double(x0, ty, w, h - (ty - y0) - 1), st, stroke != null ? stroke : new Color(0x808080));
                    if (!text.isEmpty() && sp != null) {
                        float tw = sp.width(text, size, cv.fonts);
                        cv.fillRect(x0 + 7.2, y0 - GBOX_UP, tw + 4, size * 1.2, Color.WHITE);
                        cv.text(text, x0 + 9.2, y0 + sp.ascent * size - GBOX_UP, sp, size, tc, 1, 0);
                    }
                }
                case "Button" -> {
                    RoundRectangle2D rr = new RoundRectangle2D.Double(x0, y0, w, h, 2, 2);
                    cv.fill(rr, new Color(0xF0F0F0));
                    cv.stroke(rr, new PdfCanvas.Stroke(0.75f), new Color(0x707070));
                    if (!text.isEmpty() && sp != null) {
                        float tw = sp.width(text, size, cv.fonts);
                        cv.text(text, x0 + (w - tw) / 2, y0 + h / 2 + (sp.ascent - sp.descent) * size / 2, sp, size, tc, 1, 0);
                    }
                }
                case "Drop" -> {
                    // ドロップダウン: 選んだ項目の文字と、右のボタン (灰色の地に黒い三角)
                    FontManager.Spec lf = listFont != null ? listFont : sp;
                    double bw = Math.min(w, dx * 0.75 * L.xScale);
                    cv.fillRect(x0, y0, w, h, Color.WHITE);
                    cv.fillRect(x1 - bw, y0, bw, h, BUTTON);
                    double tcx = x1 - bw / 2, tcy = y0 + h / 2;
                    cv.fill(tri(tcx - 4.8, tcy - 2.4, tcx + 4.8, tcy - 2.4, tcx, tcy + 2.6), Color.BLACK);
                    cv.stroke(new Rectangle2D.Double(x0, y0, w, h), new PdfCanvas.Stroke(0.4f), new Color(0xABABAB));
                    if (lf != null && sel >= 1 && sel <= items.size()) {
                        cv.save();
                        cv.clipRect(x0, y0, w - bw, h);
                        cv.text(items.get(sel - 1), x0 + CTL_PAD, y0 + h / 2 + (lf.ascent - lf.descent) * CTL_FONT / 2, lf, CTL_FONT, Color.BLACK, 1, 0);
                        cv.restore();
                    }
                }
                case "List" -> {
                    // リストボックス: 項目を上から並べ、選んだ項目を黒い枠で囲む。右にスクロールの列
                    FontManager.Spec lf = listFont != null ? listFont : sp;
                    double bw = Math.min(w, dx * 0.75 * L.xScale), pitch = CTL_FONT * LIST_PITCH;
                    cv.fillRect(x0, y0, w, h, Color.WHITE);
                    int visible = (int) Math.max(1, Math.floor(h / pitch));
                    int first = Math.max(0, Math.min(val, Math.max(0, items.size() - visible)));
                    cv.save();
                    cv.clipRect(x0, y0, w - bw, h);
                    for (int k = first; k < items.size(); k++) {
                        double ty = y0 + (k - first) * pitch;
                        if (ty > y0 + h) break;
                        if (lf != null) cv.text(items.get(k), x0 + CTL_PAD, ty + pitch / 2 + (lf.ascent - lf.descent) * CTL_FONT / 2, lf, CTL_FONT, Color.BLACK, 1, 0);
                        if (k + 1 == sel) cv.stroke(new Rectangle2D.Double(x0 + 0.4, ty + 0.4, w - bw - 0.8, pitch - 1.2), new PdfCanvas.Stroke(0.75f), Color.BLACK);
                    }
                    cv.restore();
                    // スクロールの列: 上下のボタン (項目が収まれば中は白)、収まらなければ溝とつまみ
                    double sx = x1 - bw, btn = Math.min(bw, h / 2);
                    if (items.size() > visible) {
                        cv.fillRect(sx, y0, bw, h, BUTTON);
                        double track = h - 2 * btn, th = Math.max(4, track * visible / items.size());
                        double f = items.size() - visible <= 0 ? 0 : (double) first / (items.size() - visible);
                        cv.fillRect(sx, y0 + btn + f * (track - th), bw, Math.min(th, 6), new Color(0xC8C8C8));
                    } else {
                        cv.fillRect(sx, y0, bw, btn, BUTTON);
                        cv.fillRect(sx, y1 - btn, bw, btn, BUTTON);
                    }
                    cv.stroke(new Rectangle2D.Double(x0, y0, w - bw, h), new PdfCanvas.Stroke(0.5f), EDGE);
                }
                case "Scroll" -> {
                    // スクロールバー: 両端の矢印のボタン (印刷では白) の間の灰色の溝と、値の位置の線
                    // ボタンの長さ: 横は Dx、縦は Dx × 1.4 (Excel の印刷の実測: 縦のボタンは横より長い)
                    double bt = dx * 0.75 * (horiz ? L.xScale : VSCROLL_BUTTON);
                    double len = horiz ? w : h;
                    bt = Math.min(bt, len / 3);
                    double f = max == min ? 0 : Math.max(0, Math.min(1, (double) (val - min) / (max - min)));
                    if (horiz) {
                        cv.fillRect(x0 + bt, y0, w - 2 * bt, h, TRACK);
                        // 溝の上下の縁 (#AAAAAA 0.5pt)
                        cv.line(x0 + bt, y0 + 0.25, x1 - bt, y0 + 0.25, new PdfCanvas.Stroke(0.5f), EDGE);
                        cv.line(x0 + bt, y1 - 0.25, x1 - bt, y1 - 0.25, new PdfCanvas.Stroke(0.5f), EDGE);
                        double px = x0 + bt + f * (w - 2 * bt);
                        cv.line(px, y0, px, y1, new PdfCanvas.Stroke(0.75f), THUMB);
                    } else {
                        cv.fillRect(x0, y0 + bt, w, h - 2 * bt, TRACK);
                        cv.line(x0 + 0.25, y0 + bt, x0 + 0.25, y1 - bt, new PdfCanvas.Stroke(0.5f), EDGE);
                        cv.line(x1 - 0.25, y0 + bt, x1 - 0.25, y1 - bt, new PdfCanvas.Stroke(0.5f), EDGE);
                        double py = y0 + bt + f * (h - 2 * bt);
                        cv.line(x0, py, x1, py, new PdfCanvas.Stroke(0.75f), THUMB);
                    }
                }
                case "Spin" -> {
                    // スピンボタン: 灰色の地に、上下 (横向きは左右) の黒い三角と、間の線
                    cv.fillRect(x0, y0, w, h, BUTTON);
                    double s = Math.max(1.5, Math.min(horiz ? h : w, (horiz ? w : h) / 2) * 0.25);
                    if (horiz) {
                        double cy = y0 + h / 2, ax = x0 + w / 4, bx2 = x1 - w / 4;
                        cv.fill(tri(ax - s / 2, cy, ax + s / 2, cy - s, ax + s / 2, cy + s), Color.BLACK);
                        cv.fill(tri(bx2 + s / 2, cy, bx2 - s / 2, cy - s, bx2 - s / 2, cy + s), Color.BLACK);
                        cv.line(x0 + w / 2, y0, x0 + w / 2, y1, new PdfCanvas.Stroke(0.4f), new Color(0xA0A0A0));
                    } else {
                        double cx = x0 + w / 2, ay = y0 + h / 4, by2 = y1 - h / 4;
                        cv.fill(tri(cx - s, ay + s / 2, cx + s, ay + s / 2, cx, ay - s / 2), Color.BLACK);
                        cv.fill(tri(cx - s, by2 - s / 2, cx + s, by2 - s / 2, cx, by2 + s / 2), Color.BLACK);
                        cv.line(x0, y0 + h / 2, x1, y0 + h / 2, new PdfCanvas.Stroke(0.4f), new Color(0xA0A0A0));
                    }
                }
                default -> { }
            }
        }
    }

    private static boolean isFalse(String v) {
        return v.equals("f") || v.equals("false");
    }

    /** VML の font の色 (auto は黒、#RRGGBB、数字は Excel の色番号)。 */
    private static Color fontColor(String v) {
        if (v == null || v.isEmpty() || v.equals("auto")) return Color.BLACK;
        try {
            if (v.startsWith("#")) return vmlColor(v);
            int idx = Integer.parseInt(v.trim());
            byte[] rgb = org.apache.poi.xssf.usermodel.DefaultIndexedColorMap.getDefaultRGB(idx);
            if (rgb != null) return new Color(rgb[0] & 255, rgb[1] & 255, rgb[2] & 255);
        } catch (RuntimeException ignored) { }
        Color c = vmlColor(v);
        return c != null ? c : Color.BLACK;
    }

    private static int parseInt(String v, int def) {
        try {
            return (int) Math.round(Double.parseDouble(v.trim()));
        } catch (NumberFormatException e) {
            return def;
        }
    }

    /** セルの範囲 (例 "$H$1:$H$6"、"Sheet2!A1:A5") のセルの表示文字列 (数式は保存された値)。 */
    static List<String> rangeTexts(XSSFSheet sheet, String ref) {
        return rangeTexts(sheet, ref, 0);
    }

    /** depth: 定義名をたどった回数 (自分自身や互いを参照する定義名で止まるよう、8 回まで)。 */
    private static List<String> rangeTexts(XSSFSheet sheet, String ref, int depth) {
        List<String> out = new ArrayList<>();
        try {
            var wb = sheet.getWorkbook();
            String r = ref.trim();
            org.apache.poi.ss.usermodel.Sheet sh = sheet;
            int bang = r.lastIndexOf('!');
            if (bang > 0) {
                String name = r.substring(0, bang).replaceAll("^'|'$", "").replace("''", "'");
                sh = wb.getSheet(name);
                r = r.substring(bang + 1);
            } else if (wb.getName(r) != null && wb.getName(r).getRefersToFormula() != null) {
                return depth < 8 ? rangeTexts(sheet, wb.getName(r).getRefersToFormula(), depth + 1) : out;
            }
            if (sh == null) return out;
            var area = new org.apache.poi.ss.util.AreaReference(r, org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
            var df = new org.apache.poi.ss.usermodel.DataFormatter();
            var first = area.getFirstCell();
            var last = area.getLastCell();
            for (int row = first.getRow(); row <= last.getRow() && out.size() < 10000; row++) {
                for (int col = first.getCol(); col <= last.getCol(); col++) {
                    var rw = sh.getRow(row);
                    var cell = rw == null ? null : rw.getCell(col);
                    if (cell == null) { out.add(""); continue; }
                    if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA) {
                        out.add(switch (cell.getCachedFormulaResultType()) {
                            case STRING -> cell.getStringCellValue();
                            case NUMERIC -> df.formatRawCellContents(cell.getNumericCellValue(), cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString());
                            case BOOLEAN -> cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                            default -> "";
                        });
                    } else out.add(df.formatCellValue(cell));
                }
            }
        } catch (Exception e) {
            // 読めない範囲は項目なし
        }
        return out;
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
            // ドロップダウン・リストボックスの項目は Windows の UI の書体 (英字 Tahoma、かな・漢字 MS UI Gothic)
            FontManager.Spec listFont = ctx.japanese ? ctx.fonts.spec("Meiryo UI", "Meiryo UI", false, false) : font;
            boolean asDisplayed = sheet.getCTWorksheet().isSetPageSetup() && sheet.getCTWorksheet().getPageSetup().isSetCellComments()
                    && sheet.getCTWorksheet().getPageSetup().getCellComments() == org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellComments.AS_DISPLAYED;
            NodeList shapes = root.getElementsByTagNameNS("*", "shape");
            for (int i = 0; i < shapes.getLength(); i++) {
              try {
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
                c.listFont = listFont;
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
                        case "Checked" -> { c.checked = !v.equals("0"); c.checkState = parseInt(v, 1); }
                        case "PrintObject" -> c.print = !v.equalsIgnoreCase("False");
                        case "NoThreeD" -> c.noThreeD = true;
                        case "Sel" -> c.sel = parseInt(v, 0);
                        case "Val" -> c.val = parseInt(v, 0);
                        case "Min" -> c.min = parseInt(v, 0);
                        case "Max" -> c.max = parseInt(v, 100);
                        case "Horiz" -> c.horiz = true;
                        case "Dx" -> c.dx = parseInt(v, 16);
                        case "FmlaRange" -> c.items = rangeTexts(sheet, v);
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
                if (tb.getLength() > 0) {
                    c.text = tb.item(0).getTextContent().replaceAll("\\s+", " ").trim();
                    // 文字の書式 (<font face size color>、size は 1/20pt)。太字は <b>
                    NodeList fl = ((Element) tb.item(0)).getElementsByTagNameNS("*", "font");
                    if (fl.getLength() > 0) {
                        Element fe = (Element) fl.item(0);
                        boolean bold = ((Element) tb.item(0)).getElementsByTagNameNS("*", "b").getLength() > 0;
                        String face = fe.getAttribute("face");
                        if (!face.isEmpty()) c.fontSpec = ctx.fonts.spec(face, bold, false);
                        else if (bold) c.fontSpec = ctx.fonts.spec(ctx.japanese ? "ＭＳ Ｐゴシック" : "Tahoma", true, false);
                        try {
                            if (!fe.getAttribute("size").isEmpty()) c.textSize = Float.parseFloat(fe.getAttribute("size")) / 20f;
                        } catch (NumberFormatException ignored) { }
                        c.textColor = fontColor(fe.getAttribute("color"));
                    }
                }
                // 塗りと線 (filled・stroked が f でなく、色の指定があるとき)
                if (type.equals("Checkbox") || type.equals("Radio") || type.equals("Label")) {
                    if (!isFalse(s.getAttribute("filled")) && !s.getAttribute("fillcolor").isEmpty()) c.fill = vmlColor(s.getAttribute("fillcolor"));
                    if (!isFalse(s.getAttribute("stroked")) && !s.getAttribute("strokecolor").isEmpty() && !type.equals("Label"))
                        c.stroke = vmlColor(s.getAttribute("strokecolor"));
                }
                if (type.equals("GBox") && !s.getAttribute("strokecolor").isEmpty()) c.stroke = vmlColor(s.getAttribute("strokecolor"));
                if (type.equals("Note")) {
                    // コメントの文字はコメントのパーツにある (VML の textbox は空のことが多い)
                    Element rw = Xml.child(cd, "Row"), cl = Xml.child(cd, "Column");
                    if (rw != null && cl != null) {
                        c.noteRow = Integer.parseInt(rw.getTextContent().trim());
                        c.noteCol = Integer.parseInt(cl.getTextContent().trim());
                        var cm = sheet.getCellComment(new org.apache.poi.ss.util.CellAddress(
                                Integer.parseInt(rw.getTextContent().trim()), Integer.parseInt(cl.getTextContent().trim())));
                        if (cm != null && cm.getString() != null) {
                            c.text = cm.getString().getString();
                            var rs = (org.apache.poi.xssf.usermodel.XSSFRichTextString) cm.getString();
                            if (rs.numFormattingRuns() > 0 && rs.getFontOfFormattingRun(0) != null) {
                                var nf = rs.getFontOfFormattingRun(0);
                                c.noteSize = nf.getFontHeightInPoints();
                                // コメントの既定の書体名「MS P ゴシック」(半角) は Excel では ＭＳ ゴシック で描かれる
                                String fn = nf.getFontName() == null ? "" : nf.getFontName();
                                if (fn.startsWith("MS P ") || fn.startsWith("MS ")) fn = "ＭＳ ゴシック";
                                if (!fn.isEmpty()) c.fontSpec = ctx.fonts.spec(fn, nf.getBold(), nf.getItalic());
                            }
                        }
                    }
                }
                out.add(c);
              } catch (IOException | RuntimeException | StackOverflowError e) {
                // 1 つのコントロールが読めなくても (画像の部品を開けない場合も)、ほかのコントロールは読む
                Warnings.add("フォームコントロールの定義を読めませんでした", e);
              }
            }
        } catch (Exception e) {
            // 読めない描画は無視する
            Warnings.add("フォームコントロールの定義を読めませんでした", e);
        }
        return out;
    }
}
