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
                    anchors.addAll(Xml.resolveAlternate(root, Set.of("a14", "x14", "cx1", "cx2", "cx4", "sle15", "sle", "tsle")));
                } catch (Exception e) {
                    // 読めない描画は無視する
                    Warnings.add("図形の定義を読めませんでした", e);
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
        /** オブジェクトが graphicFrame (グラフ・SmartArt など) か。 */
        boolean frame;
    }

    /** アンカー要素を解析する。 */
    private Anchor anchor(Element a) {
        Anchor an = new Anchor();
        an.kind = a.getLocalName();
        Element ob = objectOf(a);
        an.frame = ob != null && "graphicFrame".equals(ob.getLocalName());
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

    /**
     * セルの中のオフセット (EMU) の pt。図形・画像は、Excel と同じく画面の画素 (914400 / 画面 DPI EMU) に切り捨てる
     * (外部の報告: oneCellAnchor の colOff 3.68px・rowOff 11.64px は 3px・11px の位置に描かれる。excel-local の図形の試験 44 件 0.9739 → 0.9768)。
     * グラフなどの graphicFrame は切り捨てない (切り捨てると試験の上下が同じくらい出る。グラフの規則は別)。
     */
    private double pxFloor(long off) {
        if (curFrame) return off / EMU;
        double emuPx = 914400.0 / L.screenDpi();
        return Math.floor(off / emuPx + 1e-9) * emuPx / EMU;
    }

    /** 列 c からオフセット (EMU) の位置 (シートの pt。オフセットは列幅を超えない)。 */
    private double colPos(int c, long off) {
        return L.x(c) + Math.min(pxFloor(off) * L.xScale, L.cw(c));
    }

    /** 行 r からオフセット (EMU) の位置 (シートの pt。オフセットは行高を超えない)。 */
    private double rowPos(int r, long off) {
        return L.y(r) + Math.min(pxFloor(off) * L.yScale, L.rh(r));
    }

    /** アンカーの矩形 (シートの pt)。 */
    /** 今計算しているアンカーのオブジェクトが graphicFrame か ({@link #pxFloor})。 */
    private boolean curFrame;

    private Rectangle2D rect(Anchor a) {
        curFrame = a.frame;
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
                    List<Element> r = Xml.resolveAlternate(anchor, Set.of("a14", "cx1", "cx2", "cx4", "sle15", "sle", "tsle"));
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
    /** 右から左のシート (本体を左右反転して描く)。 */
    public boolean rtl;
    /** 白黒印刷 (pageSetup の blackAndWhite): 図形の塗りは白、線と文字は黒、グラフの系列は黒・白・灰色の順。 */
    public boolean bw;

    public void paint(PdfCanvas cv, double x0, double y0, double x1, double y1) throws IOException {
        Rectangle2D view = new Rectangle2D.Double(x0, y0, x1 - x0, y1 - y0);
        for (Element a : anchors) {
            if (!printable(a)) continue;
            Anchor an = anchor(a);
            Rectangle2D r = rect(an);
            if (!r.intersects(view) && !(r.getWidth() == 0 || r.getHeight() == 0)) continue;
            Element obj = objectOf(a);
            ctx.checkDeadline();
            try {
                if (rtl) {
                    // 右から左のシート: 位置は左右反転したまま、中身の反転を打ち消す
                    cv.save();
                    cv.transform(PdfCanvas.flipAbout(r.getCenterX()));
                }
                try {
                    paintObject(cv, obj, unrotatedBounds(obj, r), null);
                } finally {
                    if (rtl) cv.restore();
                }
            } catch (RuntimeException e) {
                // 対応していないオブジェクトでページの描画を止めない
                Warnings.add("図形を描けませんでした", e);
            }
        }
        for (VmlControls.Control c : controls) {
            if (!c.print) continue;
            try {
                if (rtl) {
                    Rectangle2D cr = c.rect(L);
                    cv.save();
                    cv.transform(PdfCanvas.flipAbout(cr.getCenterX()));
                }
                try {
                    c.paint(cv, this, L);
                } finally {
                    if (rtl) cv.restore();
                }
            } catch (RuntimeException e) {
                Warnings.add("フォームコントロールを描けませんでした", e);
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
                Warnings.add("グループ内の図形を描けませんでした", e);
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
                // DOM の要素をそのまま渡すと、XMLBeans は文書の要素として読み、中身 (gdLst・pathLst) が空になる。
                // setLoadReplaceDocumentElement(null) で、要素を型の中身として読ませる
                org.apache.xmlbeans.XmlOptions o = new org.apache.xmlbeans.XmlOptions();
                o.setLoadReplaceDocumentElement(null);
                CTCustomGeometry2D ct = CTCustomGeometry2D.Factory.parse(cust, o);
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
            Path2D path = finiteOnly(new Path2D.Double(p.getPath(c), pt));
            out.add(new GeoPath(path, p.isFilled(), p.isStroked(), p.getFill()));
        }
        return out;
    }

    /**
     * 座標が NaN の区間を除いたパス。POI は半径 0 の円弧 (調整値 0 の角丸など) を NaN にするが、半径 0 の円弧は
     * 点なので、除いても形は変わらない。
     */
    private static Path2D finiteOnly(Path2D p) {
        double[] q = new double[6];
        boolean bad = false;
        for (java.awt.geom.PathIterator it = p.getPathIterator(null); !it.isDone() && !bad; it.next()) {
            int t = it.currentSegment(q);
            int n = t == java.awt.geom.PathIterator.SEG_CLOSE ? 0 : t == java.awt.geom.PathIterator.SEG_QUADTO ? 4 : t == java.awt.geom.PathIterator.SEG_CUBICTO ? 6 : 2;
            for (int i = 0; i < n; i++) if (!Double.isFinite(q[i])) bad = true;
        }
        if (!bad) return p;
        Path2D out = new Path2D.Double(p.getWindingRule());
        boolean started = false;
        for (java.awt.geom.PathIterator it = p.getPathIterator(null); !it.isDone(); it.next()) {
            int t = it.currentSegment(q);
            int n = t == java.awt.geom.PathIterator.SEG_CLOSE ? 0 : t == java.awt.geom.PathIterator.SEG_QUADTO ? 4 : t == java.awt.geom.PathIterator.SEG_CUBICTO ? 6 : 2;
            boolean ok = true;
            for (int i = 0; i < n; i++) if (!Double.isFinite(q[i])) ok = false;
            if (!ok) continue;
            switch (t) {
                case java.awt.geom.PathIterator.SEG_MOVETO -> { out.moveTo(q[0], q[1]); started = true; }
                case java.awt.geom.PathIterator.SEG_LINETO -> { if (started) out.lineTo(q[0], q[1]); else { out.moveTo(q[0], q[1]); started = true; } }
                case java.awt.geom.PathIterator.SEG_QUADTO -> { if (started) out.quadTo(q[0], q[1], q[2], q[3]); }
                case java.awt.geom.PathIterator.SEG_CUBICTO -> { if (started) out.curveTo(q[0], q[1], q[2], q[3], q[4], q[5]); }
                default -> { if (started) out.closePath(); }
            }
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
        /** 網掛けの模様 (8 × 8、ビット r × 8 + c が 1 なら前景色)。網掛けでなければ null。 */
        Long pattern;
        /** 網掛けの前景色と背景色。 */
        Color patternFg, patternBg;
    }

    /** 網掛けの塗り (模様の名前が表になければ近似の色の単色)。 */
    static Paint pattern(String prst, Color fg, Color bg, Color approx) {
        Paint p = new Paint();
        p.color = approx;
        Long bits = prst == null ? null : PATTERNS.get(prst);
        if (bits != null) { p.pattern = bits; p.patternFg = fg; p.patternBg = bg; }
        return p;
    }

    /**
     * Excel が印刷する図形の網掛けの模様 (8 × 8。用紙上で 1 マス 0.75pt)。Excel の PDF から取り出した。
     * 倍率によらず同じ大きさで、格子は用紙の左上から「6pt の倍数 − 余白」の位置に揃う。
     */
    private static final java.util.Map<String, Long> PATTERNS = java.util.Map.ofEntries(
            Map.entry("pct5", 0x0000001000000001L),
            Map.entry("pct10", 0x0010000100100001L),
            Map.entry("pct20", 0x0044001100440011L),
            Map.entry("pct25", 0x4411441144114411L),
            Map.entry("pct30", 0x8855225588552255L),
            Map.entry("pct40", 0xA855AA558A55AA55L),
            Map.entry("pct50", 0xAA55AA55AA55AA55L),
            Map.entry("pct60", 0xAADDAA77AADDAA77L),
            Map.entry("pct70", 0xBBEEBBEEBBEEBBEEL),
            Map.entry("pct75", 0xFFBBFFEEFFBBFFEEL),
            Map.entry("pct80", 0xFF7FFFF7FF7FFFF7L),
            Map.entry("pct90", 0xFEFFFFFFEFFFFFFFL),
            Map.entry("horz", 0x00000000000000FFL),
            Map.entry("vert", 0x0101010101010101L),
            Map.entry("ltHorz", 0x000000FF000000FFL),
            Map.entry("ltVert", 0x1111111111111111L),
            Map.entry("dkHorz", 0x0000FFFF0000FFFFL),
            Map.entry("dkVert", 0x3333333333333333L),
            Map.entry("narHorz", 0x00FF00FF00FF00FFL),
            Map.entry("narVert", 0xAAAAAAAAAAAAAAAAL),
            Map.entry("dashHorz", 0x000000F00000000FL),
            Map.entry("dashVert", 0x1010101001010101L),
            Map.entry("cross", 0x01010101010101FFL),
            Map.entry("dnDiag", 0x8040201008040201L),
            Map.entry("upDiag", 0x0102040810204080L),
            Map.entry("ltDnDiag", 0x8844221188442211L),
            Map.entry("ltUpDiag", 0x1122448811224488L),
            Map.entry("dkDnDiag", 0x99CC663399CC6633L),
            Map.entry("dkUpDiag", 0x993366CC993366CCL),
            Map.entry("wdDnDiag", 0xC1E070381C0E0783L),
            Map.entry("wdUpDiag", 0x83070E1C3870E0C1L),
            Map.entry("dashDnDiag", 0x0000884422110000L),
            Map.entry("dashUpDiag", 0x0000112244880000L),
            Map.entry("diagCross", 0x8142241818244281L),
            Map.entry("smCheck", 0x9966669999666699L),
            Map.entry("lgCheck", 0xF0F0F0F00F0F0F0FL),
            Map.entry("smGrid", 0x111111FF111111FFL),
            Map.entry("lgGrid", 0x01010101010101FFL),
            Map.entry("dotGrid", 0x0001000100010055L),
            Map.entry("smConfetti", 0x2004800840021001L),
            Map.entry("lgConfetti", 0xB130031BD8C00C8DL),
            Map.entry("horzBrick", 0x101010FF010101FFL),
            Map.entry("diagBrick", 0x8142241810204080L),
            Map.entry("solidDmnd", 0x00081C3E7F3E1C08L),
            Map.entry("openDmnd", 0x8041221408142241L),
            Map.entry("dotDmnd", 0x0044001000440001L),
            Map.entry("plaid", 0x0F0F0F0FAA55AA55L),
            Map.entry("sphere", 0x1F1F19EEF1F191EEL),
            Map.entry("weave", 0x8A442811A2442A11L),
            Map.entry("divot", 0x0180010008100800L),
            Map.entry("shingle", 0x808040300C1221C0L),
            Map.entry("wave", 0x03A4180003A41800L),
            Map.entry("trellis", 0x99FF66FF99FF66FFL),
            Map.entry("zigZag", 0x1824428118244281L));

    /** 網掛けの格子の基準 (用紙の左上からの左余白・上余白、pt)。 */
    public double patternLeft, patternTop;

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
                Long bits = PATTERNS.get(prst);
                if (bits != null) {
                    p.pattern = bits;
                    p.patternFg = fg;
                    p.patternBg = bg;
                }
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

    /** SVG の拡張 (asvg:svgBlip) の中身。色の調整がある画像、白黒印刷、SVG がない画像は null (代わりの画像を描く)。 */
    private byte[] svgData(Element blip) {
        if (blip == null || drawingPart == null || bw) return null;
        for (Element c : Xml.children(blip)) {
            switch (c.getLocalName()) {
                case "grayscl", "lum", "biLevel", "clrChange", "duotone", "alphaModFix" -> { return null; }
                default -> { }
            }
        }
        for (Element ext : Xml.children(Xml.child(blip, "extLst"))) {
            for (Element sb : Xml.children(ext)) {
                if (!"svgBlip".equals(sb.getLocalName())) continue;
                String rid = sb.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
                if (rid.isEmpty()) continue;
                try {
                    PackageRelationship rel = drawingPart.getRelationship(rid);
                    if (rel == null) return null;
                    org.apache.poi.openxml4j.opc.PackagePart part = drawingPart.getRelatedPart(rel);
                    try (java.io.InputStream in = part.getInputStream()) {
                        byte[] b = in.readNBytes(32 << 20);
                        return Images.isGzip(b) ? Images.gunzip(b) : b;
                    }
                } catch (Exception e) {
                    return null;
                }
            }
        }
        return null;
    }

    /** blip の画像を埋め込む。 */
    PDImageXObject image(Element blip, double wPt, double hPt) {
        if (blip == null || drawingPart == null) return null;
        String rid = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
        if (rid.isEmpty()) rid = blip.getAttribute("r:embed");
        boolean adjust = false;
        for (Element c : Xml.children(blip)) {
            switch (c.getLocalName()) {
                case "grayscl", "lum", "biLevel", "clrChange", "duotone" -> adjust = true;
                default -> { }
            }
        }
        if (!adjust) return imageByRel(drawingPart, rid, wPt, hPt);
        return adjustedImage(blip, rid, wPt, hPt);
    }

    /**
     * 色を調整した画像 (グレースケール、明るさ・コントラスト、白黒、透明にする色、2 色)。調整は書かれた順に当てる。
     */
    private PDImageXObject adjustedImage(Element blip, String rid, double wPt, double hPt) {
        try {
            PackageRelationship rel = drawingPart.getRelationship(rid);
            if (rel == null) return null;
            PackagePart ip = drawingPart.getRelatedPart(rel);
            byte[] data;
            try (InputStream in = ip.getInputStream()) {
                data = in.readAllBytes();
            }
            String name = ip.getPartName().getName();
            StringBuilder sig = new StringBuilder(name).append('#');
            for (Element c : Xml.children(blip)) {
                sig.append(c.getLocalName());
                var attrs = c.getAttributes();
                for (int i = 0; i < attrs.getLength(); i++) sig.append(attrs.item(i).getNodeName()).append(attrs.item(i).getNodeValue());
                Color cc = ctx.theme.colorChild(c, null);
                if (cc != null) sig.append(cc.getRGB());
                for (Element d : Xml.children(c)) { Color x = ctx.theme.colorChild(d, null); if (x != null) sig.append(x.getRGB()); }
            }
            Object key = sig.toString() + "@" + Math.round(wPt) + "x" + Math.round(hPt);
            PDImageXObject cached = ctx.images.get(key);
            if (cached != null) return cached;
            java.awt.image.BufferedImage bi = Images.isMetafile(data, name) ? Images.renderMetafile(data, name, wPt, hPt, ctx.fonts) : Images.decodeLimited(data);
            if (bi == null) return null;
            java.awt.image.BufferedImage out = new java.awt.image.BufferedImage(bi.getWidth(), bi.getHeight(), java.awt.image.BufferedImage.TYPE_INT_ARGB);
            out.getGraphics().drawImage(bi, 0, 0, null);
            for (Element c : Xml.children(blip)) adjustImage(out, c);
            PDImageXObject img = org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory.createFromImage(ctx.doc, out);
            ctx.images.put(key, img);
            return img;
        } catch (Exception e) {
            Warnings.add("画像の色を調整できませんでした", e);
            return imageByRel(drawingPart, rid, wPt, hPt);
        }
    }

    /** 1 つの色の調整を画像に当てる。 */
    private void adjustImage(java.awt.image.BufferedImage img, Element c) {
        String n = c.getLocalName();
        int w = img.getWidth(), h = img.getHeight();
        double bright = Xml.longAttr(c, "bright", 0) / 100000.0, contrast = Xml.longAttr(c, "contrast", 0) / 100000.0;
        double k = contrast >= 0 ? 1 / Math.max(0.01, 1 - contrast) : 1 + contrast;
        double thresh = Xml.longAttr(c, "thresh", 50000) / 100000.0;
        Color from = null, to = null, d1 = null, d2 = null;
        if (n.equals("clrChange")) {
            from = ctx.theme.colorChild(Xml.child(c, "clrFrom"), null);
            to = ctx.theme.colorChild(Xml.child(c, "clrTo"), null);
            if (from == null || to == null) return;
        } else if (n.equals("duotone")) {
            List<Element> cs = Xml.children(c);
            if (cs.size() < 2) return;
            d1 = ctx.theme.resolve(cs.get(0), null);
            d2 = ctx.theme.resolve(cs.get(1), null);
            if (d1 == null || d2 == null) return;
        } else if (!n.equals("grayscl") && !n.equals("lum") && !n.equals("biLevel")) return;
        for (int y = 0; y < h; y++) {
            for (int x = 0; x < w; x++) {
                int argb = img.getRGB(x, y);
                int a = argb >>> 24, r = (argb >> 16) & 255, g = (argb >> 8) & 255, b = argb & 255;
                double l = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
                switch (n) {
                    case "grayscl" -> { int v = (int) Math.round(l * 255); r = g = b = v; }
                    case "lum" -> { r = lum(r, bright, k); g = lum(g, bright, k); b = lum(b, bright, k); }
                    case "biLevel" -> { int v = l >= thresh ? 255 : 0; r = g = b = v; }
                    case "clrChange" -> {
                        if (Math.abs(r - from.getRed()) <= 2 && Math.abs(g - from.getGreen()) <= 2 && Math.abs(b - from.getBlue()) <= 2) {
                            r = to.getRed(); g = to.getGreen(); b = to.getBlue(); a = a * to.getAlpha() / 255;
                        }
                    }
                    default -> { // duotone
                        r = (int) Math.round(d1.getRed() + (d2.getRed() - d1.getRed()) * l);
                        g = (int) Math.round(d1.getGreen() + (d2.getGreen() - d1.getGreen()) * l);
                        b = (int) Math.round(d1.getBlue() + (d2.getBlue() - d1.getBlue()) * l);
                    }
                }
                img.setRGB(x, y, a << 24 | r << 16 | g << 8 | b);
            }
        }
    }

    /** 明るさ・コントラストを 1 つの色の成分に当てる (0〜255)。 */
    private static int lum(int v, double bright, double k) {
        double x = (v / 255.0 - 0.5) * k + 0.5 + bright;
        return (int) Math.round(Math.max(0, Math.min(1, x)) * 255);
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
            Warnings.add("画像を読めませんでした", e);
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
        if (p.pattern != null) {
            cv.fillPattern(s, p.pattern, modify(p.patternFg, mod), modify(p.patternBg, mod), 0.75, -patternLeft, -patternTop);
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
            // 端が丸・四角の線は、端の分 (線の太さ) だけ線の部分を短くし、すき間を広げる (Office の描き方。
            // そのままでは点線の点が端でつながり、実線に見える)
            boolean capped = ls.cap == 1 || ls.cap == 2;
            for (int i = 0; i < d.length; i++) {
                double v = ls.dash[i] * w;
                if (capped) v = i % 2 == 0 ? Math.max(0.01, v - w) : v + w;
                d[i] = (float) v;
            }
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
        if (bw) {
            // 白黒印刷: 塗りは白、線は黒 (Excel の印刷)
            if (fill != null) { Paint wp = new Paint(); wp.color = java.awt.Color.WHITE; fill = wp; }
            if (ls != null) ls.color = java.awt.Color.BLACK;
        }

        Scene3D.View v3 = Scene3D.of(spPr, r);
        List<GeoPath> paths = paths(geom, adj, r);
        if (v3 != null && v3.extruded() && fill != null) {
            // 奥行き (押し出し): 奥の面から手前へ、図形をずらしながら暗い色で重ねて塗る (側面)
            Element sp3d = Xml.child(spPr, "sp3d");
            Element ec = Xml.child(sp3d, "extrusionClr");
            java.awt.Color base = ec != null ? ctx.theme.colorChild(ec, null) : fill.color != null ? fill.color
                    : fill.stopColors != null && !fill.stopColors.isEmpty() ? fill.stopColors.get(0) : java.awt.Color.GRAY;
            if (base == null) base = java.awt.Color.GRAY;
            java.awt.Color side = new java.awt.Color((int) (base.getRed() * EXTRUDE_SHADE), (int) (base.getGreen() * EXTRUDE_SHADE), (int) (base.getBlue() * EXTRUDE_SHADE));
            int steps = (int) Math.min(200, Math.ceil(Math.hypot(v3.depthX(), v3.depthY()) / 0.4));
            for (int k = steps; k >= 1; k--) {
                cv.save();
                cv.transform(AffineTransform.getTranslateInstance(v3.depthX() * k / steps, v3.depthY() * k / steps));
                cv.transform(v3.at());
                applyRotFlip(cv, r, rot, flipH, flipV);
                for (GeoPath gp : paths) if (gp.fill) cv.fill(gp.path, side);
                cv.restore();
            }
        }
        cv.save();
        if (v3 != null) cv.transform(v3.at());
        applyRotFlip(cv, r, rot, flipH, flipV);
        paintEffects(cv, spPr, style, paths, r);
        Element effl = Xml.child(spPr, "effectLst");
        // 光彩 (図形の外側のぼけた帯) と反射 (下に反転した薄い像) は図形の後ろ
        paintGlow(cv, Xml.child(effl, "glow"), paths);
        paintReflection(cv, Xml.child(effl, "reflection"), paths, fill, r);
        for (GeoPath gp : paths) {
            if (gp.fill && fill != null && gp.mod != PaintStyle.PaintModifier.NONE) fillShape(cv, gp.path, fill, gp.mod);
        }
        // 面取り・内側の影・ぼかし (ソフトエッジ) は塗りの上
        paintBevel(cv, Xml.child(spPr, "sp3d"), paths);
        paintInnerShadow(cv, Xml.child(effl, "innerShdw"), paths);
        paintSoftEdge(cv, Xml.child(effl, "softEdge"), paths);
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
            if (v3 != null) cv.transform(v3.at());
            // 文字は図形と一緒に回転するが、反転はしない
            applyRotFlip(cv, r, rot + (flipV ? 180 : 0), false, false);
            new ShapeText(ctx, this).paint(cv, tx, tr, fontRefColor, fontRefFace);
            cv.restore();
        }
    }

    /** 押し出しの側面の明るさ (塗りの色の倍。Excel の印刷では側面はほぼ黒)。 */
    private static final double EXTRUDE_SHADE = 0.3;

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

    /** 外側の影を描く (距離と方向だけずらし、ぼかし blurRad と透明度を段階の塗りで近似する)。 */
    private void paintShadow(PdfCanvas cv, Element sh, List<GeoPath> paths, Color ph) throws IOException {
        if (sh == null) return;
        Color c = ctx.theme.colorChild(sh, ph);
        if (c == null) return;
        double dist = Xml.longAttr(sh, "dist", 0) / EMU;
        double dir = Math.toRadians(Xml.longAttr(sh, "dir", 0) / 60000.0);
        double blur = Xml.longAttr(sh, "blurRad", 0) / EMU;
        if (dist <= 0 && blur <= 0) return;
        AffineTransform at = AffineTransform.getTranslateInstance(Math.cos(dir) * dist, Math.sin(dir) * dist);
        for (GeoPath gp : paths) {
            if (gp.fill) blurFill(cv, at.createTransformedShape(gp.path), c, blur);
        }
    }

    /** 文字の影の縁の濃さ (影の色の不透明度に対する割合)。exp51 wordart で 0.15・0.3・0.5 を試し 0.3 が最も合う。 */
    static final double TEXT_SHADOW_EDGE = 0.3;

    /** 文字の外側の影 (rPr の effectLst の outerShdw): 形 s をずらして、ぼかして塗る。 */
    void textShadow(PdfCanvas cv, Element sh, java.awt.Shape s) throws IOException {
        if (sh == null || s == null) return;
        Color c = ctx.theme.colorChild(sh, null);
        if (c == null) return;
        double dist = Xml.longAttr(sh, "dist", 0) / EMU;
        double dir = Math.toRadians(Xml.longAttr(sh, "dir", 0) / 60000.0);
        double blur = Xml.longAttr(sh, "blurRad", 0) / EMU;
        if (dist <= 0 && blur <= 0) return;
        java.awt.Shape m = AffineTransform.getTranslateInstance(Math.cos(dir) * dist, Math.sin(dir) * dist).createTransformedShape(s);
        double a = c.getAlpha() / 255.0;
        if (blur <= 0.5) { cv.fill(m, c); return; }
        // 縁の外は、縁で TEXT_SHADOW_EDGE 倍の濃さから、ぼかしの半径で 0 まで薄くなる (幅を変えた線を同じ濃さで重ねる)。中は文字に隠れる
        cv.fill(m, withAlpha(c, a * TEXT_SHADOW_EDGE));
        double k = 1 - Math.pow(1 - TEXT_SHADOW_EDGE * a, 1.0 / BLUR_STEPS);
        for (int i = BLUR_STEPS; i >= 1; i--) cv.stroke(m, roundStroke(2 * blur * i / BLUR_STEPS), withAlpha(c, k));
    }

    /** 影のぼかしのガウスの σ (blurRad に対する割合) と、ソフトエッジの σ (rad に対する割合)。exp64 の実測に合わせる。 */
    static final double BLUR_SIGMA = 0.5, SOFT_SIGMA = 0.35;
    /** ソフトエッジで縮めない幅 (pt)。形を rad - SOFT_KEEP 縮めて σ = SOFT_SIGMA × rad でぼかす (exp64 の 5pt・25pt の断面に合う)。 */
    static final double SOFT_KEEP = 1.7;
    /** ぼかしの画像の解像度 (画素/pt) と画素数の上限。 */
    static final double BLUR_RES = 2.0;
    static final long BLUR_MAX_PIXELS = 4_000_000;

    /** ぼかした型 (0〜1) と、その位置 (pt)。 */
    record Raster(float[] v, int w, int h, double x, double y, double res) {
        /** 型の値 → 不透明度 (0〜1) の関数 f で、色 c の画像として描く。 */
        void draw(PdfCanvas cv, Color c, java.util.function.DoubleUnaryOperator f) throws IOException {
            java.awt.image.BufferedImage img = new java.awt.image.BufferedImage(w, h, java.awt.image.BufferedImage.TYPE_INT_ARGB);
            int rgb = c.getRGB() & 0xFFFFFF;
            double a0 = c.getAlpha() / 255.0;
            int[] px = new int[w * h];
            for (int i = 0; i < px.length; i++) {
                int a = (int) Math.round(Math.max(0, Math.min(1, f.applyAsDouble(v[i]))) * a0 * 255);
                px[i] = a << 24 | rgb;
            }
            img.setRGB(0, 0, w, h, px, 0, w);
            PDImageXObject xo = org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory.createFromImage(cv.doc, img);
            cv.image(xo, x, y, w / res, h / res, 1);
        }
    }

    /** 形 s の型を σ (pt) でぼかす (箱型を 3 回)。大きすぎれば null。 */
    static Raster mask(java.awt.Shape s, double sigma) { return mask(s, sigma, 0); }

    /** 形 s を erode (pt) だけ内へ縮めた型を σ (pt) でぼかす。 */
    static Raster mask(java.awt.Shape s, double sigma, double erode) {
        Rectangle2D b = s.getBounds2D();
        double pad = Math.max(1, sigma * 3);
        double x = b.getX() - pad, y = b.getY() - pad, wPt = b.getWidth() + 2 * pad, hPt = b.getHeight() + 2 * pad;
        double res = BLUR_RES;
        while ((long) (wPt * res) * (long) (hPt * res) > BLUR_MAX_PIXELS && res > 0.25) res /= 2;
        int w = (int) Math.ceil(wPt * res), h = (int) Math.ceil(hPt * res);
        if (w <= 0 || h <= 0 || (long) w * h > BLUR_MAX_PIXELS) return null;
        java.awt.image.BufferedImage m = new java.awt.image.BufferedImage(w, h, java.awt.image.BufferedImage.TYPE_BYTE_GRAY);
        java.awt.Graphics2D g = m.createGraphics();
        g.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING, java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.scale(res, res);
        g.translate(-x, -y);
        g.fill(s);
        g.dispose();
        float[] v = new float[w * h];
        byte[] d = ((java.awt.image.DataBufferByte) m.getRaster().getDataBuffer()).getData();
        for (int i = 0; i < v.length; i++) v[i] = (d[i] & 255) / 255f;
        if (erode > 0) erode(v, w, h, erode * res);
        // σ の箱型 3 回 (幅 ≈ √(12σ²/3 + 1))
        int r = (int) Math.round((Math.sqrt(4 * sigma * res * sigma * res + 1) - 1) / 2);
        if (r > 0) for (int k = 0; k < 3; k++) { boxH(v, w, h, r); boxV(v, w, h, r); }
        return new Raster(v, w, h, x, y, res);
    }

    /** 型を e 画素だけ内へ縮める (外からの距離を 3-4 の距離変換で求め、境目は 1 画素でなめらかに)。 */
    private static void erode(float[] v, int w, int h, double e) {
        float big = 1e9f;
        float[] dist = new float[w * h];
        for (int i = 0; i < dist.length; i++) dist[i] = v[i] >= 0.5f ? big : 0;
        for (int y = 0; y < h; y++) for (int x = 0; x < w; x++) {
            int i = y * w + x;
            if (dist[i] == 0) continue;
            float m = dist[i];
            if (x > 0) m = Math.min(m, dist[i - 1] + 3);
            if (y > 0) m = Math.min(m, dist[i - w] + 3);
            if (x > 0 && y > 0) m = Math.min(m, dist[i - w - 1] + 4);
            if (x < w - 1 && y > 0) m = Math.min(m, dist[i - w + 1] + 4);
            if (x == 0 || y == 0 || x == w - 1 || y == h - 1) m = Math.min(m, 3);
            dist[i] = m;
        }
        for (int y = h - 1; y >= 0; y--) for (int x = w - 1; x >= 0; x--) {
            int i = y * w + x;
            if (dist[i] == 0) continue;
            float m = dist[i];
            if (x < w - 1) m = Math.min(m, dist[i + 1] + 3);
            if (y < h - 1) m = Math.min(m, dist[i + w] + 3);
            if (x < w - 1 && y < h - 1) m = Math.min(m, dist[i + w + 1] + 4);
            if (x > 0 && y < h - 1) m = Math.min(m, dist[i + w - 1] + 4);
            dist[i] = m;
        }
        for (int i = 0; i < v.length; i++) {
            double di = dist[i] / 3.0 - 0.5;   // 縁からの距離 (画素)
            v[i] = (float) Math.max(0, Math.min(1, di - e + 0.5));
        }
    }

    private static void boxH(float[] v, int w, int h, int r) {
        float[] row = new float[w];
        for (int yy = 0; yy < h; yy++) {
            double sum = 0;
            int o = yy * w;
            for (int i = -r; i <= r; i++) sum += v[o + Math.max(0, Math.min(w - 1, i))];
            for (int xx = 0; xx < w; xx++) {
                row[xx] = (float) (sum / (2 * r + 1));
                sum += v[o + Math.min(w - 1, xx + r + 1)] - v[o + Math.max(0, xx - r)];
            }
            System.arraycopy(row, 0, v, o, w);
        }
    }

    private static void boxV(float[] v, int w, int h, int r) {
        float[] col = new float[h];
        for (int xx = 0; xx < w; xx++) {
            double sum = 0;
            for (int i = -r; i <= r; i++) sum += v[Math.max(0, Math.min(h - 1, i)) * w + xx];
            for (int yy = 0; yy < h; yy++) {
                col[yy] = (float) (sum / (2 * r + 1));
                sum += v[Math.min(h - 1, yy + r + 1) * w + xx] - v[Math.max(0, yy - r) * w + xx];
            }
            for (int yy = 0; yy < h; yy++) v[yy * w + xx] = col[yy];
        }
    }

    /** 段階の数 (ぼかしの近似)。 */
    private static final int BLUR_STEPS = 8;

    /** 図形を、縁を blur の幅でぼかして塗る (中心から外へ薄くなる線を重ねる)。 */
    private static void blurFill(PdfCanvas cv, java.awt.Shape s, Color c, double blur) throws IOException {
        double a = c.getAlpha() / 255.0;
        if (blur <= 0.5) { cv.fill(s, c); return; }
        // Excel と同じく画像にしてぼかす (形の型をガウスぼかしにして、色の透明度にする)
        Raster m = mask(s, blur * BLUR_SIGMA);
        if (m != null) { m.draw(cv, c, v -> v); return; }
        // 中は元の濃さ。縁は幅 blur で 0 まで薄くなる (線を内側に向かって細く重ねる)
        cv.save();
        cv.fill(s, withAlpha(c, a * 0.5));
        for (int i = BLUR_STEPS; i >= 1; i--) {
            double w = blur * i / BLUR_STEPS;
            cv.stroke(s, new PdfCanvas.Stroke((float) w), withAlpha(c, a / BLUR_STEPS * 0.9));
        }
        cv.fill(s, withAlpha(c, a * 0.5));
        cv.restore();
    }

    private static Color withAlpha(Color c, double a) {
        return new Color(c.getRed(), c.getGreen(), c.getBlue(), (int) Math.max(0, Math.min(255, Math.round(a * 255))));
    }

    /** 光彩: 図形の外側に半径 rad の色の帯 (外へ薄くなる)。 */
    private void paintGlow(PdfCanvas cv, Element g, List<GeoPath> paths) throws IOException {
        if (g == null) return;
        Color c = ctx.theme.colorChild(g, null);
        double rad = Xml.longAttr(g, "rad", 0) / EMU;
        if (c == null || rad <= 0) return;
        double a = c.getAlpha() / 255.0;
        for (GeoPath gp : paths) {
            if (!gp.fill) continue;
            for (int i = BLUR_STEPS; i >= 1; i--) {
                double w = 2 * rad * i / BLUR_STEPS;
                cv.stroke(gp.path, roundStroke(w), withAlpha(c, Math.min(1, a * 1.6 / BLUR_STEPS)));
            }
        }
    }

    private static PdfCanvas.Stroke roundStroke(double w) {
        PdfCanvas.Stroke st = new PdfCanvas.Stroke((float) w);
        st.join = 1;
        st.cap = 1;
        return st;
    }

    /** ぼかし (ソフトエッジ): 図形の内側の縁を幅 rad で白に溶かす (背景を白とみなす)。 */
    private void paintSoftEdge(PdfCanvas cv, Element se, List<GeoPath> paths) throws IOException {
        if (se == null) return;
        double rad = Xml.longAttr(se, "rad", 0) / EMU;
        if (rad <= 0) return;
        for (GeoPath gp : paths) {
            if (!gp.fill) continue;
            // 画像で: 形を内へ rad - SOFT_KEEP 縮めてぼかした型を塗りの濃さとし、残りに白を重ねる。形の外には重ねない
            Raster m = mask(gp.path, rad * SOFT_SIGMA, Math.max(0, rad - SOFT_KEEP));
            if (m != null) {
                cv.save();
                // 塗りの縁のぼかし (アンチエイリアス) も消すため、形より 0.75pt 広く切る
                java.awt.geom.Area ca = new java.awt.geom.Area(gp.path);
                ca.add(new java.awt.geom.Area(new java.awt.BasicStroke(1.5f).createStrokedShape(gp.path)));
                cv.clip(ca);
                m.draw(cv, Color.WHITE, v -> 1 - v);
                cv.restore();
                continue;
            }
            // 外側は白い背景なので切り抜かずに重ねる (縁のにじみも消える)
            for (int i = BLUR_STEPS; i >= 1; i--) cv.stroke(gp.path, roundStroke(2 * rad * i / BLUR_STEPS), withAlpha(Color.WHITE, 1.8 / BLUR_STEPS));
            cv.stroke(gp.path, roundStroke(rad), Color.WHITE);
        }
    }

    /** 内側の影: 図形の中で、ずらした縁から内へぼける影。 */
    private void paintInnerShadow(PdfCanvas cv, Element sh, List<GeoPath> paths) throws IOException {
        if (sh == null) return;
        Color c = ctx.theme.colorChild(sh, null);
        if (c == null) return;
        double dist = Xml.longAttr(sh, "dist", 0) / EMU, blur = Xml.longAttr(sh, "blurRad", 0) / EMU;
        double dir = Math.toRadians(Xml.longAttr(sh, "dir", 0) / 60000.0);
        AffineTransform at = AffineTransform.getTranslateInstance(Math.cos(dir) * dist, Math.sin(dir) * dist);
        double a = c.getAlpha() / 255.0;
        for (GeoPath gp : paths) {
            if (!gp.fill) continue;
            cv.save();
            cv.clip(gp.path);
            // 図形の外側を塗った影を、ずらして内側に落とす
            java.awt.geom.Area outside = new java.awt.geom.Area(gp.path.getBounds2D().createUnion(at.createTransformedShape(gp.path).getBounds2D()));
            java.awt.geom.Rectangle2D big = gp.path.getBounds2D();
            java.awt.geom.Area ring = new java.awt.geom.Area(new Rectangle2D.Double(big.getX() - 50, big.getY() - 50, big.getWidth() + 100, big.getHeight() + 100));
            ring.subtract(new java.awt.geom.Area(at.createTransformedShape(gp.path)));
            for (int i = BLUR_STEPS; i >= 1; i--) cv.stroke(at.createTransformedShape(gp.path), roundStroke(blur * i / BLUR_STEPS), withAlpha(c, a * 0.4 / BLUR_STEPS));
            cv.fill(ring, withAlpha(c, a * 0.5));
            cv.restore();
            if (outside.isEmpty()) continue;
        }
    }

    /** 面取り (sp3d/bevelT): 上と左の縁を明るく、下と右の縁を暗くする。 */
    private void paintBevel(PdfCanvas cv, Element sp3d, List<GeoPath> paths) throws IOException {
        Element bt = Xml.child(sp3d, "bevelT");
        if (bt == null) return;
        double w = Xml.longAttr(bt, "w", 76200) / EMU;
        if (w <= 0) return;
        for (GeoPath gp : paths) {
            if (!gp.fill) continue;
            cv.save();
            cv.clip(gp.path);
            AffineTransform dn = AffineTransform.getTranslateInstance(w * 0.5, w * 0.5), up = AffineTransform.getTranslateInstance(-w * 0.5, -w * 0.5);
            cv.stroke(dn.createTransformedShape(gp.path), roundStroke(w * 1.2), withAlpha(Color.WHITE, 0.35));
            cv.stroke(up.createTransformedShape(gp.path), roundStroke(w * 1.2), withAlpha(Color.BLACK, 0.25));
            cv.restore();
        }
    }

    /** 反射: 図形を下の辺で上下に反転し、上から下へ薄くなる像を描く (stA から endA まで、endPos の位置で終わる)。 */
    private void paintReflection(PdfCanvas cv, Element rf, List<GeoPath> paths, Paint fill, Rectangle2D r) throws IOException {
        if (rf == null || fill == null || fill.color == null) return;
        double stA = Xml.longAttr(rf, "stA", 100000) / 100000.0, endA = Xml.longAttr(rf, "endA", 0) / 100000.0;
        double endPos = Xml.longAttr(rf, "endPos", 100000) / 100000.0, dist = Xml.longAttr(rf, "dist", 0) / EMU;
        AffineTransform flip = new AffineTransform(1, 0, 0, -1, 0, 2 * r.getMaxY() + dist);
        double y0 = r.getMaxY() + dist, h = r.getHeight() * endPos;
        int n = 16;
        for (GeoPath gp : paths) {
            if (!gp.fill) continue;
            java.awt.Shape m = flip.createTransformedShape(gp.path);
            for (int k = 0; k < n; k++) {
                double a = stA + (endA - stA) * (k + 0.5) / n;
                cv.save();
                cv.clipRect(r.getX() - 10, y0 + h * k / n, r.getWidth() + 20, h / n + 0.05);
                cv.fill(m, withAlpha(fill.color, a * fill.color.getAlpha() / 255.0));
                cv.restore();
            }
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
        byte[] svg = alpha < 1 ? null : svgData(blip);
        if (svg != null) {
            // SVG の画像: ベクターのまま描く (切り取りは画像全体の矩形を広げて枠で切る)。読めなければ代わりの画像
            cv.save();
            if (!isRect && !paths.isEmpty()) cv.clip(paths.get(0).path);
            Rectangle2D full = r;
            if (src != null) {
                double fw = 1 - src[0] - src[2], fh = 1 - src[1] - src[3];
                if (fw > 0 && fh > 0) {
                    cv.clipRect(r.getX(), r.getY(), r.getWidth(), r.getHeight());
                    full = new Rectangle2D.Double(r.getX() - src[0] * r.getWidth() / fw, r.getY() - src[1] * r.getHeight() / fh,
                            r.getWidth() / fw, r.getHeight() / fh);
                }
            }
            boolean ok = SvgPainter.paint(ctx, cv, svg, full);
            cv.restore();
            if (ok) img = null;
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
            if ("timeslicer".equals(c.getLocalName())) {
                try {
                    new TimelinePainter(ctx, sheet).paint(cv, c.getAttribute("name"), r);
                } catch (Exception e) {
                    Warnings.add("タイムラインを描けませんでした", e);
                }
                continue;
            }
            if ("slicer".equals(c.getLocalName())) {
                // スライサー: 定義はシートの slicers とブックの slicerCache にある
                try {
                    new SlicerPainter(ctx, sheet).paint(cv, c.getAttribute("name"), r);
                } catch (Exception e) {
                    Warnings.add("スライサーを描けませんでした", e);
                }
                continue;
            }
            String rid = c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            if (rid.isEmpty()) rid = c.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "dm");
            if (rid.isEmpty()) continue;
            try {
                PackageRelationship rel = drawingPart.getRelationship(rid);
                if (rel == null) continue;
                PackagePart part = drawingPart.getRelatedPart(rel);
                if ("chart".equals(c.getLocalName()) && part.getPartName().getName().toLowerCase(java.util.Locale.ROOT).contains("chartex")) {
                    new ChartExPainter(ctx, part).paint(cv, r);
                } else if ("chart".equals(c.getLocalName())) {
                    new ChartPainter(ctx, this, part).paint(cv, r);
                } else if ("relIds".equals(c.getLocalName())) {
                    paintDiagram(cv, part, r);
                }
            } catch (Exception e) {
                // 対応していないフレーム
                Warnings.add("グラフまたは SmartArt を描けませんでした", e);
            }
        }
    }

    /** SmartArt を描く: Excel はデータモデルの隣に描画済みの図形 (dsp) を保存している。 */
    private void paintDiagram(PdfCanvas cv, PackagePart dataPart, Rectangle2D r) throws Exception {
        // この SmartArt の描画: データの dsp:dataModelExt の relId (シートの描画からの関係) が指す部分。
        // 見つからなければ、最初の diagramDrawing を使う (古いファイル)
        String own = null;
        if (dataPart != null) {
            try (InputStream in = dataPart.getInputStream()) {
                java.util.regex.Matcher m = java.util.regex.Pattern.compile("dataModelExt[^>]*relId=\"([^\"]+)\"")
                        .matcher(new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8));
                if (m.find()) own = m.group(1);
            }
        }
        for (PackageRelationship rel : drawingPart.getRelationships()) {
            if (!rel.getRelationshipType().endsWith("/diagramDrawing")) continue;
            if (own != null && !own.equals(rel.getId()) && drawingPart.getRelationship(own) != null) continue;
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
