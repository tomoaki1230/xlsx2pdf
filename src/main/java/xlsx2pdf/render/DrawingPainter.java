package xlsx2pdf.render;

import java.awt.Color;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.sl.draw.geom.Context;
import org.apache.poi.sl.draw.geom.CustomGeometry;
import org.apache.poi.sl.draw.geom.Guide;
import org.apache.poi.sl.draw.geom.GuideIf;
import org.apache.poi.sl.draw.geom.PathIf;
import org.apache.poi.sl.draw.geom.PresetGeometries;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.xslf.draw.geom.XSLFCustomGeometry;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.main.CTCustomGeometry2D;
import org.w3c.dom.Element;

/** DrawingML のオブジェクト (図形、画像、テキストボックス、グループ、コネクタ、グラフ) と VML のフォームコントロールを描く。 */
public final class DrawingPainter {
    /** 1 pt あたりの EMU。 */
    static final double EMU = 12700.0;

    private final RenderContext ctx;
    private final XSSFSheet sheet;
    /** レイアウト (グラフシート・図形の範囲の事前調査では null)。 */
    private SheetLayout L;
    /** 描画パーツ (なければ null)。 */
    private PackagePart drawingPart;
    /** 描画のアンカー要素 (twoCellAnchor など)。 */
    private final List<Element> anchors = new ArrayList<>();
    /** VML のフォームコントロール。 */
    private final List<VmlControls.Control> controls;

    /**
     * シートの描画とフォームコントロールを読み込む。
     *
     * @param layout レイアウト (図形の範囲の事前調査やグラフシートでは null)
     */
    public DrawingPainter(RenderContext ctx, XSSFSheet sheet, SheetLayout layout) {
        this.ctx = ctx;
        this.sheet = sheet;
        this.L = layout;
        for (POIXMLDocumentPart p : sheet.getRelations()) {
            if (p instanceof XSSFDrawing d) {
                drawingPart = d.getPackagePart();
                try (InputStream in = drawingPart.getInputStream()) {
                    Element root = Xml.parse(in);
                    anchors.addAll(Xml.resolveAlternate(root, Set.of("a14", "x14")));
                } catch (Exception e) {
                    // 読めない描画は無視する
                }
                break;
            }
        }
        this.controls = VmlControls.load(ctx, sheet);
    }

    // ------------------------------------------------------------ アンカー

    /** 解析したアンカー。 */
    private static final class Anchor {
        /** 開始・終了のセル (0 始まりの行・列)。 */
        int r0 = -1, c0 = -1, r1 = -1, c1 = -1;
        /** セル内オフセット、絶対位置、大きさ (EMU)。 */
        long r0Off, c0Off, r1Off, c1Off, x, y, cx, cy;
        /** アンカーの種類 (twoCellAnchor, oneCellAnchor, absoluteAnchor)。 */
        String kind;
    }

    /** アンカー要素を解析する。 */
    private Anchor anchor(Element a) {
        Anchor an = new Anchor();
        an.kind = a.getLocalName();
        Element from = Xml.child(a, "from"), to = Xml.child(a, "to");
        if (from != null) {
            an.c0 = (int) Xml.longAttr(Xml.child(from, "col"), "_", 0);
            an.c0 = Integer.parseInt(Xml.text(Xml.child(from, "col")).trim());
            an.c0Off = Long.parseLong(Xml.text(Xml.child(from, "colOff")).trim());
            an.r0 = Integer.parseInt(Xml.text(Xml.child(from, "row")).trim());
            an.r0Off = Long.parseLong(Xml.text(Xml.child(from, "rowOff")).trim());
        }
        if (to != null) {
            an.c1 = Integer.parseInt(Xml.text(Xml.child(to, "col")).trim());
            an.c1Off = Long.parseLong(Xml.text(Xml.child(to, "colOff")).trim());
            an.r1 = Integer.parseInt(Xml.text(Xml.child(to, "row")).trim());
            an.r1Off = Long.parseLong(Xml.text(Xml.child(to, "rowOff")).trim());
        }
        Element pos = Xml.child(a, "pos"), ext = Xml.child(a, "ext");
        if (pos != null) { an.x = Xml.longAttr(pos, "x", 0); an.y = Xml.longAttr(pos, "y", 0); }
        if (ext != null) { an.cx = Xml.longAttr(ext, "cx", 0); an.cy = Xml.longAttr(ext, "cy", 0); }
        return an;
    }

    /** 列 c からオフセット (EMU) の位置 (シートの pt。オフセットは列幅を超えない)。 */
    private double colPos(int c, long off) {
        return L.x(c) + Math.min(off / EMU * L.xScale, L.cw(c));
    }

    /** 行 r からオフセット (EMU) の位置 (シートの pt。オフセットは行高を超えない)。 */
    private double rowPos(int r, long off) {
        return L.y(r) + Math.min(off / EMU * L.yScale, L.rh(r));
    }

    /** アンカーの矩形 (シートの pt)。 */
    private Rectangle2D rect(Anchor a) {
        switch (a.kind) {
            case "twoCellAnchor" -> {
                double x0 = colPos(a.c0, a.c0Off), y0 = rowPos(a.r0, a.r0Off);
                double x1 = colPos(a.c1, a.c1Off), y1 = rowPos(a.r1, a.r1Off);
                return new Rectangle2D.Double(x0, y0, Math.max(0, x1 - x0), Math.max(0, y1 - y0));
            }
            case "oneCellAnchor" -> {
                double x0 = colPos(a.c0, a.c0Off), y0 = rowPos(a.r0, a.r0Off);
                // 1 セルアンカーのオブジェクトの大きさは絶対値: Excel はセルと一緒に縮めない
                return new Rectangle2D.Double(x0, y0, a.cx / EMU, a.cy / EMU);
            }
            default -> {
                return new Rectangle2D.Double(a.x / EMU * L.xScale, a.y / EMU * L.yScale, a.cx / EMU, a.cy / EMU);
            }
        }
    }

    /** 印刷するオブジェクトが覆う範囲 {最終行, 最終列, 先頭行, 先頭列}。なければ null。 */
    public int[] extent() {
        int lr = -1, lc = -1, fr = Integer.MAX_VALUE, fc = Integer.MAX_VALUE;
        SheetLayout tmp = null;
        for (Element a : anchors) {
            if (!printable(a)) continue;
            Anchor an = anchor(a);
            if (an.kind.equals("twoCellAnchor")) {
                // 高さや幅が 0 の図形も、始まりの行・列は使用範囲に入る
                // 作成者の PC の PDF では数えない
                boolean zero = !ctx.printer.authorEnv;
                lr = Math.max(lr, Math.max(zero ? an.r0 : -1, an.r1Off > 0 ? an.r1 : an.r1 - 1));
                lc = Math.max(lc, Math.max(zero ? an.c0 : -1, an.c1Off > 0 ? an.c1 : an.c1 - 1));
                fr = Math.min(fr, an.r0);
                fc = Math.min(fc, an.c0);
            } else {
                if (tmp == null) tmp = new SheetLayout(sheet, ctx.mdw, ctx.printer, 600, 600, 3000, 300);
                SheetLayout saved = L;
                L = tmp;
                Rectangle2D r = rect(an);
                L = saved;
                int[] end = cellAt(tmp, r.getMaxX(), r.getMaxY());
                int[] st = cellAt(tmp, r.getX(), r.getY());
                lr = Math.max(lr, end[0]);
                lc = Math.max(lc, end[1]);
                fr = Math.min(fr, st[0]);
                fc = Math.min(fc, st[1]);
            }
        }
        for (VmlControls.Control c : controls) {
            if (!c.print) continue;
            lr = Math.max(lr, c.r1);
            lc = Math.max(lc, c.c1);
            fr = Math.min(fr, c.r0);
            fc = Math.min(fc, c.c0);
        }
        if (lr < 0) return null;
        return new int[] {lr, lc, fr, fc};
    }

    /** 印刷するオブジェクトまたはコントロールが、セルのブロックに重なるか (現在の倍率のシートの pt で判定)。 */
    public boolean hasObjectIn(int r0, int c0, int r1, int c1) {
        double x0 = L.x(c0), y0 = L.y(r0), x1 = L.x(c1 + 1), y1 = L.y(r1 + 1);
        for (Element a : anchors) {
            if (!printable(a)) continue;
            Rectangle2D r = rect(anchor(a));
            if (r.getMaxX() > x0 && r.getX() < x1 && r.getMaxY() > y0 && r.getY() < y1) return true;
        }
        for (VmlControls.Control c : controls) {
            if (c.print && c.r1 >= r0 && c.r0 <= r1 && c.c1 >= c0 && c.c0 <= c1) return true;
        }
        return false;
    }

    /** 位置 (x, y) を含むセル {行, 列}。 */
    private static int[] cellAt(SheetLayout l, double x, double y) {
        int c = 0, r = 0;
        while (c < l.maxCol - 1 && l.x(c + 1) < x) c++;
        while (r < l.maxRow - 1 && l.y(r + 1) < y) r++;
        return new int[] {r, c};
    }

    /** 印刷するオブジェクトか (非表示・印刷しない設定を除く)。 */
    private boolean printable(Element anchor) {
        Element cd = Xml.child(anchor, "clientData");
        if (cd != null && cd.hasAttribute("fPrintsWithSheet") && !Xml.boolAttr(cd, "fPrintsWithSheet", true)) return false;
        Element obj = objectOf(anchor);
        if (obj == null) return false;
        Element nv = nvPr(obj);
        return nv == null || !Xml.boolAttr(nv, "hidden", false);
    }

    /** アンカーの中のオブジェクト要素 (sp, pic, grpSp, graphicFrame, cxnSp)。 */
    private static Element objectOf(Element anchor) {
        for (Element c : Xml.children(anchor)) {
            switch (c.getLocalName()) {
                case "sp", "grpSp", "pic", "graphicFrame", "cxnSp" -> { return c; }
                case "AlternateContent" -> {
                    List<Element> r = Xml.resolveAlternate(anchor, Set.of("a14"));
                    for (Element x : r) {
                        String n = x.getLocalName();
                        if (n.equals("sp") || n.equals("grpSp") || n.equals("pic") || n.equals("graphicFrame") || n.equals("cxnSp")) return x;
                    }
                }
                default -> { }
            }
        }
        return null;
    }

    /** オブジェクトの非表示プロパティ (cNvPr)。 */
    private static Element nvPr(Element obj) {
        for (Element c : Xml.children(obj)) {
            if (c.getLocalName().startsWith("nv")) return Xml.child(c, "cNvPr");
        }
        return null;
    }

    // ------------------------------------------------------------ 描画

    /**
     * すべてのオブジェクトを描く (クリップは見えるブロックに設定済み)。
     *
     * @param x0 ブロックの左端 (シートの pt)
     * @param y0 ブロックの上端
     * @param x1 ブロックの右端
     * @param y1 ブロックの下端
     */
    public void paint(PdfCanvas cv, double x0, double y0, double x1, double y1) throws IOException {
        Rectangle2D view = new Rectangle2D.Double(x0, y0, x1 - x0, y1 - y0);
        for (Element a : anchors) {
            if (!printable(a)) continue;
            Anchor an = anchor(a);
            Rectangle2D r = rect(an);
            if (!r.intersects(view) && !(r.getWidth() == 0 || r.getHeight() == 0)) continue;
            Element obj = objectOf(a);
            try {
                paintObject(cv, obj, unrotatedBounds(obj, r), null);
            } catch (RuntimeException e) {
                // 対応していないオブジェクトでページの描画を止めない
            }
        }
        for (VmlControls.Control c : controls) {
            if (!c.print) continue;
            try {
                c.paint(cv, this, L);
            } catch (RuntimeException e) {
                // 無視する
            }
        }
    }

    /** グラフシートのグラフを矩形に描く。 */
    public void paintChartSheet(PdfCanvas cv, double x, double y, double w, double h) throws IOException {
        for (Element a : anchors) {
            Element obj = objectOf(a);
            if (obj == null) continue;
            Anchor an = anchor(a);
            Rectangle2D r = an.kind.equals("absoluteAnchor") && an.cx > 0
                    ? new Rectangle2D.Double(x, y, w, h)
                    : new Rectangle2D.Double(x, y, w, h);
            paintObject(cv, obj, r, null);
        }
    }

    /**
     * オブジェクトを矩形 r に描く。グループの子は、子の xfrm とグループの変換から r を求める。
     *
     * @param parentGroupFill グループの塗りつぶし (grpFill の参照先。なければ null)
     */
    private void paintObject(PdfCanvas cv, Element obj, Rectangle2D r, Element parentGroupFill) throws IOException {
        switch (obj.getLocalName()) {
            case "sp", "cxnSp" -> paintShape(cv, obj, r, parentGroupFill);
            case "pic" -> paintPicture(cv, obj, r);
            case "grpSp" -> paintGroup(cv, obj, r);
            case "graphicFrame" -> paintFrame(cv, obj, r);
            default -> { }
        }
    }

    /** グループを描く (子の座標をグループの座標系から変換する)。 */
    private void paintGroup(PdfCanvas cv, Element g, Rectangle2D r) throws IOException {
        Element gp = Xml.child(g, "grpSpPr");
        Element xfrm = Xml.child(gp, "xfrm");
        double chX = 0, chY = 0, chW = 1, chH = 1;
        double rot = 0;
        boolean flipH = false, flipV = false;
        if (xfrm != null) {
            Element chOff = Xml.child(xfrm, "chOff"), chExt = Xml.child(xfrm, "chExt");
            Element ext = Xml.child(xfrm, "ext");
            chX = Xml.longAttr(chOff, "x", 0) / EMU;
            chY = Xml.longAttr(chOff, "y", 0) / EMU;
            chW = Xml.longAttr(chExt, "cx", Xml.longAttr(ext, "cx", 1)) / EMU;
            chH = Xml.longAttr(chExt, "cy", Xml.longAttr(ext, "cy", 1)) / EMU;
            rot = Xml.longAttr(xfrm, "rot", 0) / 60000.0;
            flipH = Xml.boolAttr(xfrm, "flipH", false);
            flipV = Xml.boolAttr(xfrm, "flipV", false);
        }
        if (chW == 0) chW = 1;
        if (chH == 0) chH = 1;
        double sx = r.getWidth() / chW, sy = r.getHeight() / chH;
        cv.save();
        applyRotFlip(cv, r, rot, flipH, flipV);
        Element grpFill = gp == null ? null : fillElement(gp);
        for (Element child : Xml.resolveAlternate(g, Set.of("a14"))) {
            String n = child.getLocalName();
            if (!(n.equals("sp") || n.equals("grpSp") || n.equals("pic") || n.equals("graphicFrame") || n.equals("cxnSp")))
                continue;
            Element nv = nvPr(child);
            if (nv != null && Xml.boolAttr(nv, "hidden", false)) continue;
            Element cx = childXfrm(child);
            if (cx == null) continue;
            Element off = Xml.child(cx, "off"), ext = Xml.child(cx, "ext");
            double x = Xml.longAttr(off, "x", 0) / EMU, y = Xml.longAttr(off, "y", 0) / EMU;
            double w = Xml.longAttr(ext, "cx", 0) / EMU, h = Xml.longAttr(ext, "cy", 0) / EMU;
            Rectangle2D cr = new Rectangle2D.Double(r.getX() + (x - chX) * sx, r.getY() + (y - chY) * sy, w * sx, h * sy);
            try {
                paintObject(cv, child, cr, grpFill);
            } catch (RuntimeException e) {
                // 飛ばす
            }
        }
        cv.restore();
    }

    /**
     * セルアンカーは回転後のオブジェクトの外接矩形を保存する。90 / 270 度に近い回転では Excel は幅と高さを入れ替えるので、
     * 回転前の図形は、同じ中心で辺を入れ替えた矩形になる。
     */
    static Rectangle2D unrotatedBounds(Element obj, Rectangle2D r) {
        Element x = obj == null ? null : childXfrm(obj);
        if (x == null) return r;
        double rot = ((Xml.longAttr(x, "rot", 0) / 60000.0) % 360 + 360) % 360;
        boolean swap = (rot >= 45 && rot < 135) || (rot >= 225 && rot < 315);
        if (!swap) return r;
        double cx = r.getCenterX(), cy = r.getCenterY();
        return new Rectangle2D.Double(cx - r.getHeight() / 2, cy - r.getWidth() / 2, r.getHeight(), r.getWidth());
    }

    /** オブジェクトの xfrm (図形は spPr の中、グラフィックフレームは直下)。 */
    private static Element childXfrm(Element obj) {
        for (Element c : Xml.children(obj)) {
            String n = c.getLocalName();
            if (n.equals("spPr") || n.equals("grpSpPr")) return Xml.child(c, "xfrm");
            if (n.equals("xfrm")) return c; // graphicFrame の xfrm
        }
        return null;
    }

    /** 矩形の中心を基準に回転・反転を適用する。 */
    private static void applyRotFlip(PdfCanvas cv, Rectangle2D r, double rotDeg, boolean flipH, boolean flipV) throws IOException {
        if (rotDeg == 0 && !flipH && !flipV) return;
        AffineTransform at = new AffineTransform();
        at.translate(r.getCenterX(), r.getCenterY());
        if (rotDeg != 0) at.rotate(Math.toRadians(rotDeg));
        at.scale(flipH ? -1 : 1, flipV ? -1 : 1);
        at.translate(-r.getCenterX(), -r.getCenterY());
        cv.transform(at);
    }

    // ------------------------------------------------------------ 形状

    /** 図形の形状 (プリセットまたはカスタム)。 */
    static CustomGeometry geometry(Element spPr) {
        Element prst = Xml.child(spPr, "prstGeom");
        if (prst != null) {
            CustomGeometry g = PresetGeometries.getInstance().get(prst.getAttribute("prst"));
            return g != null ? g : PresetGeometries.getInstance().get("rect");
        }
        Element cust = Xml.child(spPr, "custGeom");
        if (cust != null) {
            try {
                CTCustomGeometry2D ct = CTCustomGeometry2D.Factory.parse(cust);
                return XSLFCustomGeometry.convertCustomGeometry(ct);
            } catch (Exception e) {
                return PresetGeometries.getInstance().get("rect");
            }
        }
        return PresetGeometries.getInstance().get("rect");
    }

    /** 形状の調整値 (avLst)。 */
    static Map<String, GuideIf> adjusts(Element spPr) {
        Map<String, GuideIf> m = new HashMap<>();
        Element av = Xml.path(spPr, "prstGeom", "avLst");
        if (av == null) av = Xml.path(spPr, "custGeom", "avLst");
        for (Element gd : Xml.children(av, "gd")) {
            Guide g = new Guide();
            g.setName(gd.getAttribute("name"));
            g.setFmla(gd.getAttribute("fmla"));
            m.put(g.getName(), g);
        }
        return m;
    }

    /** 形状の 1 つのパス (塗りつぶすか、線を引くか、塗りの明暗)。 */
    record GeoPath(Path2D path, boolean fill, boolean stroke, PaintStyle.PaintModifier mod) {}

    /** 形状のパスを矩形 r に合わせて求める。 */
    static List<GeoPath> paths(CustomGeometry geom, Map<String, GuideIf> adj, Rectangle2D r) {
        // POI の arcTo は、アンカーが原点でないと絶対座標とアンカー基準の座標を混ぜてしまうため、
        // (0,0,w,h) で評価してから実際の位置に動かす
        List<GeoPath> out = new ArrayList<>();
        Rectangle2D anchor = new Rectangle2D.Double(0, 0, Math.max(r.getWidth(), 1e-3), Math.max(r.getHeight(), 1e-3));
        Context c = new Context(geom, anchor, adj::get);
        AffineTransform at = AffineTransform.getTranslateInstance(r.getX(), r.getY());
        for (PathIf p : geom) {
            AffineTransform pt = at;
            if (p.getW() > 0 && p.getH() > 0) {
                // パス独自の座標系 (<path w="5" h="5">。flowChart 系の標準図形) を図形の大きさに広げる。POI はこの拡大をしない
                pt = new AffineTransform(at);
                pt.scale(anchor.getWidth() / p.getW(), anchor.getHeight() / p.getH());
            }
            Path2D path = new Path2D.Double(p.getPath(c), pt);
            out.add(new GeoPath(path, p.isFilled(), p.isStroked(), p.getFill()));
        }
        return out;
    }

    /** 形状の文字の矩形 (形状の rect 定義から)。 */
    static Rectangle2D textRect(CustomGeometry geom, Map<String, GuideIf> adj, Rectangle2D r) {
        try {
            var tb = geom.getTextBounds();
            if (tb == null) return r;
            Rectangle2D anchor = new Rectangle2D.Double(0, 0, Math.max(r.getWidth(), 1e-3), Math.max(r.getHeight(), 1e-3));
            Context c = new Context(geom, anchor, adj::get);
            Rectangle2D b = tb.getPath(c).getBounds2D();
            return new Rectangle2D.Double(r.getX() + b.getX(), r.getY() + b.getY(), b.getWidth(), b.getHeight());
        } catch (RuntimeException e) {
            return r;
        }
    }

    // ------------------------------------------------------------ 塗りつぶし・線

    /** spPr の塗りつぶし要素 (solidFill, gradFill, blipFill, pattFill, noFill, grpFill)。 */
    static Element fillElement(Element spPr) {
        for (Element c : Xml.children(spPr)) {
            switch (c.getLocalName()) {
                case "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill" -> { return c; }
                default -> { }
            }
        }
        return null;
    }

    /** 解決した塗りつぶし: 単色、グラデーション、または画像。 */
    static final class Paint {
        /** 単色。 */
        Color color;
        /** グラデーションの位置 (0〜1) と色。 */
        List<Double> stopPos;
        List<Color> stopColors;
        /** 線形グラデーションの角度 (度)。 */
        double angle;
        /** 放射状のグラデーションか。 */
        boolean radial;
        /** 図形と一緒に回転するか。 */
        boolean rotWithShape = true;
        /** 画像の塗りつぶし。 */
        PDImageXObject image;
        /** 画像の切り取り {左, 上, 右, 下} (割合)。 */
        double[] srcRect;
    }

    /**
     * 塗りつぶし要素を解決する。
     *
     * @param phClr   スタイル参照のプレースホルダー色
     * @param grpFill グループの塗りつぶし (grpFill のとき使う)
     */
    Paint paint(Element fill, Color phClr, Element grpFill) {
        if (fill == null) return null;
        Paint p = new Paint();
        switch (fill.getLocalName()) {
            case "noFill" -> { return null; }
            case "solidFill" -> {
                p.color = ctx.theme.colorChild(fill, phClr);
                if (p.color == null) return null;
            }
            case "gradFill" -> {
                p.stopPos = new ArrayList<>();
                p.stopColors = new ArrayList<>();
                for (Element gs : Xml.children(Xml.child(fill, "gsLst"), "gs")) {
                    p.stopPos.add(Xml.longAttr(gs, "pos", 0) / 100000.0);
                    Color c = ctx.theme.colorChild(gs, phClr);
                    p.stopColors.add(c == null ? Color.WHITE : c);
                }
                if (p.stopPos.isEmpty()) return null;
                // 位置の順に並べる
                for (int i = 0; i < p.stopPos.size(); i++)
                    for (int j = i + 1; j < p.stopPos.size(); j++)
                        if (p.stopPos.get(j) < p.stopPos.get(i)) {
                            double t = p.stopPos.get(i); p.stopPos.set(i, p.stopPos.get(j)); p.stopPos.set(j, t);
                            Color c = p.stopColors.get(i); p.stopColors.set(i, p.stopColors.get(j)); p.stopColors.set(j, c);
                        }
                Element lin = Xml.child(fill, "lin");
                p.angle = lin != null ? Xml.longAttr(lin, "ang", 0) / 60000.0 : 90;
                p.radial = Xml.child(fill, "path") != null;
                p.color = p.stopColors.get(0);
            }
            case "pattFill" -> {
                Color fg = ctx.theme.colorChild(Xml.child(fill, "fgClr"), phClr);
                Color bg = ctx.theme.colorChild(Xml.child(fill, "bgClr"), phClr);
                if (fg == null) fg = Color.BLACK;
                if (bg == null) bg = Color.WHITE;
                String prst = fill.getAttribute("prst");
                double cov = 0.5;
                if (prst.startsWith("pct")) {
                    try { cov = Integer.parseInt(prst.substring(3)) / 100.0; } catch (NumberFormatException ignored) { cov = 0.5; }
                } else if (prst.startsWith("lt") || prst.startsWith("narrow")) cov = 0.25;
                else if (prst.startsWith("dk") || prst.startsWith("wd")) cov = 0.6;
                p.color = CellPainter.mix(fg, bg, cov);
            }
            case "blipFill" -> {
                Element blip = Xml.child(fill, "blip");
                p.image = image(blip, 400, 400);
                if (p.image == null) return null;
                Element src = Xml.child(fill, "srcRect");
                if (src != null) p.srcRect = new double[] {Xml.longAttr(src, "l", 0) / 100000.0, Xml.longAttr(src, "t", 0) / 100000.0,
                        Xml.longAttr(src, "r", 0) / 100000.0, Xml.longAttr(src, "b", 0) / 100000.0};
            }
            case "grpFill" -> { return grpFill == null ? null : paint(grpFill, phClr, null); }
            default -> { return null; }
        }
        return p;
    }

    /** blip の画像を埋め込む。 */
    PDImageXObject image(Element blip, double wPt, double hPt) {
        if (blip == null || drawingPart == null) return null;
        String rid = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
        if (rid.isEmpty()) rid = blip.getAttribute("r:embed");
        return imageByRel(drawingPart, rid, wPt, hPt);
    }

    /** パーツの関係 ID の画像を埋め込む。 */
    PDImageXObject imageByRel(PackagePart part, String rid, double wPt, double hPt) {
        try {
            PackageRelationship rel = part.getRelationship(rid);
            if (rel == null) return null;
            PackagePart ip = part.getRelatedPart(rel);
            byte[] data;
            try (InputStream in = ip.getInputStream()) {
                data = in.readAllBytes();
            }
            String name = ip.getPartName().getName();
            boolean meta = Images.isMetafile(data, name);
            Object key = meta ? name + "@" + Math.round(wPt) + "x" + Math.round(hPt) : name;
            return Images.toPdf(ctx, key, data, name, wPt, hPt);
        } catch (Exception e) {
            return null;
        }
    }

    /** 図形を塗りつぶす (画像・グラデーション・単色)。 */
    void fillShape(PdfCanvas cv, Shape s, Paint p, PaintStyle.PaintModifier mod) throws IOException {
        if (p == null) return;
        if (p.image != null) {
            Rectangle2D b = s.getBounds2D();
            cv.save();
            cv.clip(s);
            drawCropped(cv, p.image, b, p.srcRect);
            cv.restore();
            return;
        }
        if (p.stopPos != null && p.stopPos.size() > 1) {
            fillGradient(cv, s, p, mod);
            return;
        }
        cv.fill(s, modify(p.color, mod));
    }

    /** 塗りの明暗の修飾 (暗く・明るく) を適用する。 */
    static Color modify(Color c, PaintStyle.PaintModifier mod) {
        if (c == null || mod == null) return c;
        return switch (mod) {
            case DARKEN -> scale(c, 0.6);
            case DARKEN_LESS -> scale(c, 0.8);
            case LIGHTEN -> lighten(c, 0.4);
            case LIGHTEN_LESS -> lighten(c, 0.2);
            default -> c;
        };
    }

    /** 色を暗くする (各成分 × f)。 */
    private static Color scale(Color c, double f) {
        return new Color((int) (c.getRed() * f), (int) (c.getGreen() * f), (int) (c.getBlue() * f), c.getAlpha());
    }

    /** 色を白に近づける (割合 f)。 */
    private static Color lighten(Color c, double f) {
        return new Color((int) (c.getRed() + (255 - c.getRed()) * f), (int) (c.getGreen() + (255 - c.getGreen()) * f),
                (int) (c.getBlue() + (255 - c.getBlue()) * f), c.getAlpha());
    }

    /** グラデーションを帯で近似して塗る。 */
    private void fillGradient(PdfCanvas cv, Shape s, Paint p, PaintStyle.PaintModifier mod) throws IOException {
        Rectangle2D b = s.getBounds2D();
        cv.save();
        cv.clip(s);
        int n = 64;
        if (p.radial) {
            double cx = b.getCenterX(), cy = b.getCenterY();
            double rad = Math.hypot(b.getWidth(), b.getHeight()) / 2;
            for (int i = 0; i < n; i++) {
                double t = i / (double) (n - 1);
                Color c = gradColor(p, 1 - t);
                double rr = rad * (1 - i / (double) n);
                cv.fill(new Ellipse2D.Double(cx - rr, cy - rr, rr * 2, rr * 2), modify(c, mod));
            }
        } else {
            double ang = Math.toRadians(p.angle);
            double cx = b.getCenterX(), cy = b.getCenterY();
            double dx = Math.cos(ang), dy = Math.sin(ang);
            // グラデーションの方向に沿った外接矩形の範囲
            double half = Math.abs(b.getWidth() / 2 * dx) + Math.abs(b.getHeight() / 2 * dy);
            double len = Math.hypot(b.getWidth(), b.getHeight());
            for (int i = 0; i < n; i++) {
                double t0 = i / (double) n, t1 = (i + 1) / (double) n;
                Color c = gradColor(p, (t0 + t1) / 2);
                double s0 = -half + 2 * half * t0, s1 = -half + 2 * half * t1 + 0.3;
                // 方向に沿って s0〜s1 の帯 (横方向には十分に広く)
                Path2D band = new Path2D.Double();
                double px = -dy, py = dx;
                band.moveTo(cx + dx * s0 + px * len, cy + dy * s0 + py * len);
                band.lineTo(cx + dx * s1 + px * len, cy + dy * s1 + py * len);
                band.lineTo(cx + dx * s1 - px * len, cy + dy * s1 - py * len);
                band.lineTo(cx + dx * s0 - px * len, cy + dy * s0 - py * len);
                band.closePath();
                cv.fill(band, modify(c, mod));
            }
        }
        cv.restore();
    }

    /** グラデーションの位置 t の色。 */
    static Color gradColor(Paint p, double t) {
        List<Double> pos = p.stopPos;
        List<Color> cols = p.stopColors;
        if (t <= pos.get(0)) return cols.get(0);
        for (int i = 1; i < pos.size(); i++) {
            if (t <= pos.get(i)) {
                double a = pos.get(i - 1), b = pos.get(i);
                double f = b > a ? (t - a) / (b - a) : 1;
                Color c0 = cols.get(i - 1), c1 = cols.get(i);
                return new Color((int) Math.round(c0.getRed() + (c1.getRed() - c0.getRed()) * f),
                        (int) Math.round(c0.getGreen() + (c1.getGreen() - c0.getGreen()) * f),
                        (int) Math.round(c0.getBlue() + (c1.getBlue() - c0.getBlue()) * f),
                        (int) Math.round(c0.getAlpha() + (c1.getAlpha() - c0.getAlpha()) * f));
            }
        }
        return cols.get(cols.size() - 1);
    }

    /** 線のスタイル。 */
    static final class LineStyle {
        /** 線の色。 */
        Color color;
        /** 太さ (pt)。 */
        double width;
        /** 破線 (太さの倍数。実線なら null)。 */
        float[] dash;
        /** 線端 (0 平ら、1 丸、2 四角) と角 (0 マイター、1 丸、2 面取り)。 */
        int cap = 0, join = 1;
        /** 始点・終点の矢印 (なければ null)。 */
        Element head, tail;
    }

    /** 線のスタイルを解決する (スタイルの lnRef を基に、明示の ln で上書き。線なしなら null)。 */
    LineStyle line(Element ln, Color phClr, Element styleLnRef) {
        // スタイルの線を先に解決し、明示の ln で上書きする
        LineStyle ls = null;
        if (styleLnRef != null) {
            int idx = Xml.intAttr(styleLnRef, "idx", 0);
            Color refColor = ctx.theme.colorChild(styleLnRef, phClr);
            if (idx > 0 && idx <= ctx.theme.lineStyles.size()) {
                Element tln = ctx.theme.lineStyles.get(idx - 1);
                ls = lineFrom(tln, refColor, null);
            } else if (idx > 0) {
                ls = new LineStyle();
                ls.color = refColor;
                ls.width = 0.75 * idx;
            }
        }
        if (ln != null) ls = lineFrom(ln, phClr, ls);
        if (ls == null || ls.color == null) return null;
        return ls;
    }

    /** ln 要素の指定を base に重ねる。 */
    private LineStyle lineFrom(Element ln, Color phClr, LineStyle base) {
        LineStyle ls = new LineStyle();
        if (base != null) {
            ls.color = base.color; ls.width = base.width; ls.dash = base.dash; ls.cap = base.cap; ls.join = base.join;
            ls.head = base.head; ls.tail = base.tail;
        } else {
            ls.width = 0.75;
        }
        if (ln.hasAttribute("w")) ls.width = Xml.longAttr(ln, "w", 9525) / EMU;
        String cap = ln.getAttribute("cap");
        if (cap.equals("rnd")) ls.cap = 1; else if (cap.equals("sq")) ls.cap = 2; else if (cap.equals("flat")) ls.cap = 0;
        for (Element c : Xml.children(ln)) {
            switch (c.getLocalName()) {
                case "noFill" -> ls.color = null;
                case "solidFill" -> ls.color = ctx.theme.colorChild(c, phClr);
                case "gradFill" -> {
                    Element gs = Xml.child(Xml.child(c, "gsLst"), "gs");
                    ls.color = ctx.theme.colorChild(gs, phClr);
                }
                case "prstDash" -> ls.dash = presetDash(c.getAttribute("val"));
                case "custDash" -> {
                    List<Float> d = new ArrayList<>();
                    for (Element ds : Xml.children(c, "ds")) {
                        d.add(Xml.longAttr(ds, "d", 100000) / 100000f);
                        d.add(Xml.longAttr(ds, "sp", 100000) / 100000f);
                    }
                    ls.dash = new float[d.size()];
                    for (int i = 0; i < d.size(); i++) ls.dash[i] = d.get(i);
                }
                case "round" -> ls.join = 1;
                case "bevel" -> ls.join = 2;
                case "miter" -> ls.join = 0;
                case "headEnd" -> ls.head = c;
                case "tailEnd" -> ls.tail = c;
                default -> { }
            }
        }
        return ls;
    }

    /** プリセットの破線のパターン (線の太さの倍数)。 */
    static float[] presetDash(String v) {
        return switch (v) {
            case "dot", "sysDot" -> new float[] {1, 1};
            case "dash", "sysDash" -> v.equals("sysDash") ? new float[] {3, 1} : new float[] {4, 3};
            case "lgDash" -> new float[] {8, 3};
            case "dashDot", "sysDashDot" -> v.equals("sysDashDot") ? new float[] {3, 1, 1, 1} : new float[] {4, 3, 1, 3};
            case "lgDashDot" -> new float[] {8, 3, 1, 3};
            case "lgDashDotDot", "sysDashDotDot" -> new float[] {8, 3, 1, 3, 1, 3};
            default -> null;
        };
    }

    /** 図形の輪郭を描く。 */
    void strokeShape(PdfCanvas cv, Shape s, LineStyle ls) throws IOException {
        if (ls == null || ls.color == null || ls.width <= 0) return;
        double w = Math.max(ls.width, 0.1);
        PdfCanvas.Stroke st = new PdfCanvas.Stroke((float) w);
        if (ls.dash != null) {
            float[] d = new float[ls.dash.length];
            for (int i = 0; i < d.length; i++) d[i] = (float) (ls.dash[i] * w);
            st.dash(d);
        }
        st.cap = ls.cap;
        st.join = ls.join;
        cv.stroke(s, st, ls.color);
    }

    /** 開いたパスの両端に矢印を描く。 */
    void arrows(PdfCanvas cv, Path2D path, LineStyle ls) throws IOException {
        if (ls == null || ls.color == null) return;
        List<Point2D> pts = new ArrayList<>();
        var it = path.getPathIterator(null, 0.5);
        double[] c = new double[6];
        while (!it.isDone()) {
            int t = it.currentSegment(c);
            if (t == java.awt.geom.PathIterator.SEG_MOVETO || t == java.awt.geom.PathIterator.SEG_LINETO)
                pts.add(new Point2D.Double(c[0], c[1]));
            it.next();
        }
        if (pts.size() < 2) return;
        if (ls.head != null) arrow(cv, ls.head, pts.get(0), pts.get(1), ls);
        if (ls.tail != null) arrow(cv, ls.tail, pts.get(pts.size() - 1), pts.get(pts.size() - 2), ls);
    }

    private void arrow(PdfCanvas cv, Element end, Point2D tip, Point2D from, LineStyle ls) throws IOException {
        String type = end.getAttribute("type");
        if (type.isEmpty() || type.equals("none")) return;
        // 矢じりの大きさの基準は線幅だが、細い線では 0.7mm (約 1.98pt) を下限にする
        double lw = Math.max(ls.width, 0.7 / 25.4 * 72);
        double wf = switch (end.getAttribute("w")) { case "sm" -> 2; case "lg" -> 5; default -> 3; };
        double lf = switch (end.getAttribute("len")) { case "sm" -> 2; case "lg" -> 5; default -> 3; };
        double aw = wf * lw, al = lf * lw;
        double ang = Math.atan2(tip.getY() - from.getY(), tip.getX() - from.getX());
        AffineTransform at = AffineTransform.getTranslateInstance(tip.getX(), tip.getY());
        at.rotate(ang);
        Path2D p = new Path2D.Double();
        switch (type) {
            case "oval" -> {
                Shape e = at.createTransformedShape(new Ellipse2D.Double(-al / 2, -aw / 2, al, aw));
                cv.fill(e, ls.color);
                return;
            }
            case "diamond" -> {
                p.moveTo(al / 2, 0); p.lineTo(0, -aw / 2); p.lineTo(-al / 2, 0); p.lineTo(0, aw / 2); p.closePath();
            }
            case "arrow" -> {
                p.moveTo(-al, -aw / 2); p.lineTo(0, 0); p.lineTo(-al, aw / 2);
                PdfCanvas.Stroke st = new PdfCanvas.Stroke((float) ls.width);
                st.join = 0;
                cv.stroke(at.createTransformedShape(p), st, ls.color);
                return;
            }
            case "stealth" -> {
                p.moveTo(0, 0); p.lineTo(-al, -aw / 2); p.lineTo(-al * 0.6, 0); p.lineTo(-al, aw / 2); p.closePath();
            }
            default -> { p.moveTo(0, 0); p.lineTo(-al, -aw / 2); p.lineTo(-al, aw / 2); p.closePath(); }
        }
        cv.fill(at.createTransformedShape(p), ls.color);
    }

    // ------------------------------------------------------------ 図形

    /** 図形 (sp, cxnSp) を描く: 影、塗りつぶし、線、矢印、文字。 */
    private void paintShape(PdfCanvas cv, Element sp, Rectangle2D r, Element grpFill) throws IOException {
        Element spPr = Xml.child(sp, "spPr");
        Element style = Xml.child(sp, "style");
        Element xfrm = Xml.child(spPr, "xfrm");
        double rot = xfrm != null ? Xml.longAttr(xfrm, "rot", 0) / 60000.0 : 0;
        boolean flipH = xfrm != null && Xml.boolAttr(xfrm, "flipH", false);
        boolean flipV = xfrm != null && Xml.boolAttr(xfrm, "flipV", false);
        CustomGeometry geom = geometry(spPr);
        Map<String, GuideIf> adj = adjusts(spPr);

        // 塗りつぶし
        Element fillEl = fillElement(spPr);
        Paint fill;
        if (fillEl != null) {
            fill = paint(fillEl, null, grpFill);
        } else {
            fill = styleFill(style);
        }
        Element lnRef = style == null ? null : Xml.child(style, "lnRef");
        LineStyle ls = line(Xml.child(spPr, "ln"), lnRef == null ? null : ctx.theme.colorChild(lnRef, null), lnRef);
        boolean connector = sp.getLocalName().equals("cxnSp");
        if (connector) fill = null;

        cv.save();
        applyRotFlip(cv, r, rot, flipH, flipV);
        List<GeoPath> paths = paths(geom, adj, r);
        paintEffects(cv, spPr, style, paths, r);
        for (GeoPath gp : paths) {
            if (gp.fill && fill != null && gp.mod != PaintStyle.PaintModifier.NONE) fillShape(cv, gp.path, fill, gp.mod);
        }
        for (GeoPath gp : paths) {
            if (gp.stroke && ls != null) {
                strokeShape(cv, gp.path, ls);
                if (ls.head != null || ls.tail != null) arrows(cv, gp.path, ls);
            }
        }
        cv.restore();

        Element tx = Xml.child(sp, "txBody");
        if (tx != null) {
            Rectangle2D tr = textRect(geom, adj, r);
            Color fontRefColor = null;
            String fontRefFace = null;
            if (style != null) {
                Element fr = Xml.child(style, "fontRef");
                if (fr != null) {
                    fontRefColor = ctx.theme.colorChild(fr, null);
                    fontRefFace = "major".equals(fr.getAttribute("idx")) ? "+mj-lt" : "+mn-lt";
                }
            }
            cv.save();
            // 文字は図形と一緒に回転するが、反転はしない
            applyRotFlip(cv, r, rot + (flipV ? 180 : 0), false, false);
            new ShapeText(ctx, this).paint(cv, tx, tr, fontRefColor, fontRefFace);
            cv.restore();
        }
    }

    /** スタイル参照 (fillRef) の塗りつぶし。 */
    Paint styleFill(Element style) {
        if (style == null) return null;
        Element fr = Xml.child(style, "fillRef");
        if (fr == null) return null;
        int idx = Xml.intAttr(fr, "idx", 0);
        Color c = ctx.theme.colorChild(fr, null);
        if (idx == 0 || c == null) return null;
        int i = idx >= 1001 ? idx - 1001 : idx - 1;
        if (i >= 0 && i < ctx.theme.fillStyles.size()) {
            Element f = ctx.theme.fillStyles.get(i);
            Paint p = paint(f, c, null);
            if (p != null) return p;
        }
        Paint p = new Paint();
        p.color = c;
        return p;
    }

    /** 効果 (外側の影のみ簡易に対応) を描く。 */
    private void paintEffects(PdfCanvas cv, Element spPr, Element style, List<GeoPath> paths, Rectangle2D r) throws IOException {
        Element eff = Xml.child(spPr, "effectLst");
        if (eff == null && style != null) {
            Element er = Xml.child(style, "effectRef");
            int idx = er == null ? 0 : Xml.intAttr(er, "idx", 0);
            if (idx > 0 && idx <= ctx.theme.effectStyles.size()) {
                eff = Xml.child(ctx.theme.effectStyles.get(idx - 1), "effectLst");
                if (eff != null) {
                    Color ph = ctx.theme.colorChild(er, null);
                    paintShadow(cv, Xml.child(eff, "outerShdw"), paths, ph);
                }
                return;
            }
        }
        if (eff != null) paintShadow(cv, Xml.child(eff, "outerShdw"), paths, null);
    }

    /** 外側の影を描く (距離と方向だけずらした塗りつぶし)。 */
    private void paintShadow(PdfCanvas cv, Element sh, List<GeoPath> paths, Color ph) throws IOException {
        if (sh == null) return;
        Color c = ctx.theme.colorChild(sh, ph);
        if (c == null) return;
        double dist = Xml.longAttr(sh, "dist", 0) / EMU;
        double dir = Math.toRadians(Xml.longAttr(sh, "dir", 0) / 60000.0);
        if (dist <= 0) return;
        AffineTransform at = AffineTransform.getTranslateInstance(Math.cos(dir) * dist, Math.sin(dir) * dist);
        for (GeoPath gp : paths) {
            if (gp.fill) cv.fill(at.createTransformedShape(gp.path), c);
        }
    }

    // ------------------------------------------------------------ 画像

    /** 画像 (pic) を描く: 後ろの塗りつぶし、切り取り、透明度、枠線。 */
    private void paintPicture(PdfCanvas cv, Element pic, Rectangle2D r) throws IOException {
        Element spPr = Xml.child(pic, "spPr");
        Element bf = Xml.child(pic, "blipFill");
        Element xfrm = Xml.child(spPr, "xfrm");
        double rot = xfrm != null ? Xml.longAttr(xfrm, "rot", 0) / 60000.0 : 0;
        boolean flipH = xfrm != null && Xml.boolAttr(xfrm, "flipH", false);
        boolean flipV = xfrm != null && Xml.boolAttr(xfrm, "flipV", false);
        Element blip = Xml.child(bf, "blip");
        PDImageXObject img = image(blip, r.getWidth(), r.getHeight());
        double[] src = null;
        Element sr = Xml.child(bf, "srcRect");
        if (sr != null) src = new double[] {Xml.longAttr(sr, "l", 0) / 100000.0, Xml.longAttr(sr, "t", 0) / 100000.0,
                Xml.longAttr(sr, "r", 0) / 100000.0, Xml.longAttr(sr, "b", 0) / 100000.0};
        float alpha = 1;
        Element am = blip == null ? null : Xml.child(blip, "alphaModFix");
        if (am != null) alpha = Xml.longAttr(am, "amt", 100000) / 100000f;
        CustomGeometry geom = geometry(spPr);
        Map<String, GuideIf> adj = adjusts(spPr);
        cv.save();
        applyRotFlip(cv, r, rot, flipH, flipV);
        List<GeoPath> paths = paths(geom, adj, r);
        Element prst = Xml.child(spPr, "prstGeom");
        boolean isRect = prst == null || "rect".equals(prst.getAttribute("prst"));
        // 画像の後ろの塗りつぶし (spPr の塗りつぶし)
        Element fe = fillElement(spPr);
        if (fe != null) {
            Paint p = paint(fe, null, null);
            for (GeoPath gp : paths) if (gp.fill) fillShape(cv, gp.path, p, gp.mod);
        }
        if (img != null) {
            cv.save();
            if (!isRect && !paths.isEmpty()) cv.clip(paths.get(0).path);
            if (alpha < 1) {
                drawCroppedAlpha(cv, img, r, src, alpha);
            } else {
                drawCropped(cv, img, r, src);
            }
            cv.restore();
        }
        Element style = Xml.child(pic, "style");
        Element lnRef = style == null ? null : Xml.child(style, "lnRef");
        LineStyle ls = line(Xml.child(spPr, "ln"), null, lnRef);
        if (ls != null) for (GeoPath gp : paths) if (gp.stroke) strokeShape(cv, gp.path, ls);
        cv.restore();
    }

    /** 切り取りを適用して画像を描く。 */
    static void drawCropped(PdfCanvas cv, PDImageXObject img, Rectangle2D r, double[] src) throws IOException {
        drawCroppedAlpha(cv, img, r, src, 1);
    }

    /**
     * 切り取りと透明度を適用して画像を描く。
     *
     * @param src 切り取り {左, 上, 右, 下} (割合。null なら切り取りなし)
     */
    static void drawCroppedAlpha(PdfCanvas cv, PDImageXObject img, Rectangle2D r, double[] src, float alpha) throws IOException {
        if (src == null || (src[0] == 0 && src[1] == 0 && src[2] == 0 && src[3] == 0)) {
            cv.image(img, r.getX(), r.getY(), r.getWidth(), r.getHeight(), alpha);
            return;
        }
        double fw = 1 - src[0] - src[2], fh = 1 - src[1] - src[3];
        if (fw <= 0 || fh <= 0) return;
        double fullW = r.getWidth() / fw, fullH = r.getHeight() / fh;
        cv.save();
        cv.clipRect(r.getX(), r.getY(), r.getWidth(), r.getHeight());
        cv.image(img, r.getX() - src[0] * fullW, r.getY() - src[1] * fullH, fullW, fullH, alpha);
        cv.restore();
    }

    // ------------------------------------------------------------ グラフィックフレーム (グラフ)

    /** グラフィックフレーム (グラフ、SmartArt) を描く。 */
    private void paintFrame(PdfCanvas cv, Element gf, Rectangle2D r) throws IOException {
        Element gd = Xml.path(gf, "graphic", "graphicData");
        if (gd == null) return;
        for (Element c : Xml.children(gd)) {
            String rid = c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            if (rid.isEmpty()) rid = c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "dm");
            if (rid.isEmpty()) continue;
            try {
                PackageRelationship rel = drawingPart.getRelationship(rid);
                if (rel == null) continue;
                PackagePart part = drawingPart.getRelatedPart(rel);
                if ("chart".equals(c.getLocalName())) {
                    new ChartPainter(ctx, this, part).paint(cv, r);
                } else if ("relIds".equals(c.getLocalName())) {
                    paintDiagram(cv, part, r);
                }
            } catch (Exception e) {
                // 対応していないフレーム
            }
        }
    }

    /** SmartArt を描く: Excel はデータモデルの隣に描画済みの図形 (dsp) を保存している。 */
    private void paintDiagram(PdfCanvas cv, PackagePart dataPart, Rectangle2D r) throws Exception {
        for (PackageRelationship rel : drawingPart.getRelationships()) {
            if (!rel.getRelationshipType().endsWith("/diagramDrawing")) continue;
            PackagePart dp = drawingPart.getRelatedPart(rel);
            Element root;
            try (InputStream in = dp.getInputStream()) {
                root = Xml.parse(in);
            }
            Element tree = Xml.child(root, "spTree");
            PackagePart saved = drawingPart;
            drawingPart = dp;
            try {
                for (Element sp : Xml.children(tree)) {
                    if (!sp.getLocalName().equals("sp") && !sp.getLocalName().equals("grpSp")) continue;
                    Element x = childXfrm(sp);
                    if (x == null) continue;
                    double ox = Xml.longAttr(Xml.child(x, "off"), "x", 0) / EMU, oy = Xml.longAttr(Xml.child(x, "off"), "y", 0) / EMU;
                    double w = Xml.longAttr(Xml.child(x, "ext"), "cx", 0) / EMU, h = Xml.longAttr(Xml.child(x, "ext"), "cy", 0) / EMU;
                    Rectangle2D cr = new Rectangle2D.Double(r.getX() + ox, r.getY() + oy, w, h);
                    Element txXfrm = Xml.child(sp, "txXfrm");
                    paintObject(cv, sp, cr, null);
                    if (txXfrm != null) { /* 文字は図形の矩形の中に描く */ }
                }
            } finally {
                drawingPart = saved;
            }
            return;
        }
    }

}
