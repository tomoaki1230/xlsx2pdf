package xlsx2pdf.render;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Arc2D;
import java.awt.geom.Area;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Line2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.geom.RoundRectangle2D;
import java.awt.image.BufferedImage;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayDeque;
import java.util.Deque;
import java.util.HashMap;
import java.util.Map;

/**
 * EMF (拡張メタファイル) を Java2D に描く。Excel が貼り付けた「図としてコピー」の範囲や QR コードなどは EMF で保存され、
 * POI の描画は範囲の単位の取り違えなどで崩れる。
 *
 * <p>座標は「論理 → ワールド変換 → ウィンドウ/ビューポートの対応 → デバイス (参照機器のピクセル) → 出力画像」の順に変換し、
 * 図形・文字・画像はすべて出力画像の座標に直してから描く。対応する主な記録: 座標系 (SETWINDOW*, SETVIEWPORT*, SETMAPMODE,
 * SET/MODIFYWORLDTRANSFORM, SAVEDC/RESTOREDC)、オブジェクト (ペン、ブラシ、フォント)、線と多角形 (16/32 ビット)、
 * 矩形・楕円・角丸矩形、パス、文字 (EXTTEXTOUTW/A)、画像 (BITBLT, STRETCHBLT, STRETCHDIBITS, SETDIBITSTODEVICE)、
 * クリップ (INTERSECTCLIPRECT, EXCLUDECLIPRECT, EXTSELECTCLIPRGN, SELECTCLIPPATH)。EMF+ の記録 (コメント) は使わない。
 */
final class EmfRenderer {
    private final ByteBuffer b;
    private final Graphics2D g;
    private final FontManager fonts;
    /** デバイス → 出力画像。 */
    private AffineTransform devToOut;
    /** 参照機器の 1 ピクセルの出力画像での大きさ。 */
    private double devPx = 1;
    /** 描いた記録の数 (0 なら EMF+ だけのファイルなどとして別の描画に任せる)。 */
    int drawn;

    /** 描画の状態 (SAVEDC で保存する)。 */
    private static final class State implements Cloneable {
        double wox, woy, wex = 1, wey = 1, vox, voy, vex = 1, vey = 1;
        int mapMode = 1;
        AffineTransform world = new AffineTransform();
        Pen pen = new Pen(0, 0, Color.BLACK);
        Brush brush = new Brush(0, Color.WHITE);
        LogFont font = new LogFont(-12, 0, 400, false, false, false, "ＭＳ Ｐゴシック");
        Color textColor = Color.BLACK, bkColor = Color.WHITE;
        int bkMode = 2, textAlign = 0, polyFill = 1;
        double cpx, cpy;
        Area clip;

        @Override
        public State clone() {
            try {
                State s = (State) super.clone();
                s.world = new AffineTransform(world);
                s.clip = clip == null ? null : new Area(clip);
                return s;
            } catch (CloneNotSupportedException e) {
                throw new IllegalStateException(e);
            }
        }
    }

    private record Pen(int style, double width, Color color) {}

    private record Brush(int style, Color color, java.awt.Paint paint) {
        Brush(int style, Color color) { this(style, color, null); }
    }

    private record LogFont(int height, int escapement, int weight, boolean italic, boolean underline, boolean strike, String face) {}

    private State st = new State();
    private final Deque<State> saved = new ArrayDeque<>();
    private final Map<Integer, Object> objects = new HashMap<>();
    /** 作成中のパス (出力画像の座標)。 */
    private Path2D path;
    /** パスが閉じた後に使える最後のパス。 */
    private Path2D lastPath;

    EmfRenderer(byte[] data, Graphics2D g, FontManager fonts) {
        this.b = ByteBuffer.wrap(data).order(ByteOrder.LITTLE_ENDIAN);
        this.g = g;
        this.fonts = fonts;
    }

    /** EMF を出力画像 (w × h ピクセル) いっぱいに描く。 */
    void draw(double w, double h) {
        // ヘッダー: 画像の範囲 rclFrame (0.01 mm) と、参照機器のピクセル数 szlDevice・大きさ szlMillimeters
        int fl = b.getInt(24), ft = b.getInt(28), fr = b.getInt(32), fb = b.getInt(36);
        int devW = b.getInt(72), devH = b.getInt(76), mmW = b.getInt(80), mmH = b.getInt(84);
        double pxPerMmX = mmW > 0 ? (double) devW / mmW : 96 / 25.4, pxPerMmY = mmH > 0 ? (double) devH / mmH : 96 / 25.4;
        int hdrSize = b.getInt(4);
        if (hdrSize >= 108 && b.capacity() >= 108) {
            // szlMicrometers があればより正確
            int umW = b.getInt(100), umH = b.getInt(104);
            if (umW > 0 && umH > 0) { pxPerMmX = devW * 1000.0 / umW; pxPerMmY = devH * 1000.0 / umH; }
        }
        double x0 = fl / 100.0 * pxPerMmX, y0 = ft / 100.0 * pxPerMmY;
        double fw = (fr - fl) / 100.0 * pxPerMmX, fh = (fb - ft) / 100.0 * pxPerMmY;
        if (fw <= 0 || fh <= 0) {
            // 範囲がなければ rclBounds (デバイス単位)
            x0 = b.getInt(8); y0 = b.getInt(12); fw = b.getInt(16) - x0 + 1; fh = b.getInt(20) - y0 + 1;
        }
        devToOut = new AffineTransform(w / fw, 0, 0, h / fh, -x0 * w / fw, -y0 * h / fh);
        devPx = Math.sqrt(Math.abs(devToOut.getDeterminant()));
        int pos = 0;
        while (pos + 8 <= b.capacity()) {
            int type = b.getInt(pos), size = b.getInt(pos + 4);
            if (size < 8 || pos + size > b.capacity()) break;
            try {
                record(type, pos, size);
            } catch (RuntimeException e) {
                // 壊れた記録は飛ばす
                Warnings.add("EMF 画像の一部を描けませんでした", e);
            }
            if (type == 14) break;
            pos += size;
        }
    }

    // ---------------------------------------------------------------- 座標

    /** 論理座標 → 出力画像の変換。 */
    private AffineTransform logToOut() {
        AffineTransform t = new AffineTransform(devToOut);
        double sx = 1, sy = 1;
        if (st.mapMode == 7 || st.mapMode == 8) {
            // MM_ISOTROPIC / MM_ANISOTROPIC
            sx = st.wex == 0 ? 1 : st.vex / st.wex;
            sy = st.wey == 0 ? 1 : st.vey / st.wey;
            if (st.mapMode == 7) {
                double s = Math.min(Math.abs(sx), Math.abs(sy));
                sx = Math.signum(sx) * s; sy = Math.signum(sy) * s;
            }
        } else if (st.mapMode != 1) {
            // MM_LOMETRIC 等: 論理単位の長さ (mm) × デバイスの mm あたりピクセル、y は上向き
            double mm = switch (st.mapMode) { case 2 -> 0.1; case 3 -> 0.01; case 4 -> 0.254; case 5 -> 0.0254; case 6 -> 25.4 / 1440; default -> 1; };
            double pxmm = b.getInt(72) / (double) Math.max(1, b.getInt(80));
            sx = mm * pxmm; sy = -mm * pxmm;
        }
        t.translate(st.vox, st.voy);
        t.scale(sx, sy);
        t.translate(-st.wox, -st.woy);
        t.concatenate(st.world);
        return t;
    }

    private Point2D pt(double x, double y) {
        return logToOut().transform(new Point2D.Double(x, y), null);
    }

    /** 論理単位の長さの、出力画像での大きさ (変換の平均倍率)。 */
    private double logScale() {
        return Math.sqrt(Math.abs(logToOut().getDeterminant()));
    }

    // ---------------------------------------------------------------- 記録

    private int i32(int p) { return b.getInt(p); }

    private short i16(int p) { return b.getShort(p); }

    private float f32(int p) { return b.getFloat(p); }

    private static Color color(int v) {
        return new Color(v & 255, (v >> 8) & 255, (v >> 16) & 255);
    }

    private void record(int type, int p, int size) {
        switch (type) {
            case 9 -> { st.wex = i32(p + 8); st.wey = i32(p + 12); }          // SETWINDOWEXTEX
            case 10 -> { st.wox = i32(p + 8); st.woy = i32(p + 12); }         // SETWINDOWORGEX
            case 11 -> { st.vex = i32(p + 8); st.vey = i32(p + 12); }         // SETVIEWPORTEXTEX
            case 12 -> { st.vox = i32(p + 8); st.voy = i32(p + 12); }         // SETVIEWPORTORGEX
            case 17 -> st.mapMode = i32(p + 8);                               // SETMAPMODE
            case 18 -> st.bkMode = i32(p + 8);                                // SETBKMODE
            case 19 -> st.polyFill = i32(p + 8);                              // SETPOLYFILLMODE
            case 22 -> st.textAlign = i32(p + 8);                             // SETTEXTALIGN
            case 24 -> st.textColor = color(i32(p + 8));                      // SETTEXTCOLOR
            case 25 -> st.bkColor = color(i32(p + 8));                        // SETBKCOLOR
            case 33 -> saved.push(st.clone());                                // SAVEDC
            case 34 -> {                                                      // RESTOREDC
                int n = i32(p + 8);
                for (int k = 0; k < (n < 0 ? -n : 1) && !saved.isEmpty(); k++) st = saved.pop();
                applyClip();
            }
            case 35 -> st.world = xform(p + 8);                               // SETWORLDTRANSFORM
            case 36 -> {                                                      // MODIFYWORLDTRANSFORM
                AffineTransform x = xform(p + 8);
                switch (i32(p + 32)) {
                    case 1 -> st.world = new AffineTransform();
                    case 2 -> st.world.concatenate(x);
                    case 3 -> st.world.preConcatenate(x);
                    case 4 -> st.world = x;
                    default -> { }
                }
            }
            case 37 -> select(i32(p + 8));                                    // SELECTOBJECT
            case 40 -> objects.remove(i32(p + 8));                            // DELETEOBJECT
            case 38 -> objects.put(i32(p + 8), new Pen(i32(p + 12), i32(p + 16), color(i32(p + 24)))); // CREATEPEN
            case 95 -> objects.put(i32(p + 8), new Pen(i32(p + 28), i32(p + 32), color(i32(p + 40)))); // EXTCREATEPEN
            case 39 -> objects.put(i32(p + 8), new Brush(i32(p + 12), color(i32(p + 16))));             // CREATEBRUSHINDIRECT
            case 93, 94 -> patternBrush(p);                                   // CREATEMONOBRUSH, CREATEDIBPATTERNBRUSHPT
            case 118 -> gradientFill(p);                                      // GRADIENTFILL
            case 82 -> objects.put(i32(p + 8), logFont(p + 12));              // EXTCREATEFONTINDIRECTW
            case 27 -> { st.cpx = i32(p + 8); st.cpy = i32(p + 12); if (path != null) { Point2D q = pt(st.cpx, st.cpy); path.moveTo(q.getX(), q.getY()); } } // MOVETOEX
            case 54 -> lineTo(i32(p + 8), i32(p + 12));                       // LINETO
            case 2, 3, 4, 5, 6 -> poly(type, p, false);                       // POLYBEZIER, POLYGON, POLYLINE, POLYBEZIERTO, POLYLINETO
            case 85, 86, 87, 88, 89 -> poly(type - 83, p, true);              // 16 ビット版
            case 7, 8 -> polyPoly(type == 8, p, false);                       // POLYPOLYLINE, POLYPOLYGON
            case 90, 91 -> polyPoly(type == 91, p, true);                     // 16 ビット版
            case 43 -> shape(new Rectangle2D.Double(Math.min(i32(p + 8), i32(p + 16)), Math.min(i32(p + 12), i32(p + 20)),
                    Math.abs(i32(p + 16) - i32(p + 8)), Math.abs(i32(p + 20) - i32(p + 12))), true);    // RECTANGLE
            case 42 -> shape(new Ellipse2D.Double(Math.min(i32(p + 8), i32(p + 16)), Math.min(i32(p + 12), i32(p + 20)),
                    Math.abs(i32(p + 16) - i32(p + 8)), Math.abs(i32(p + 20) - i32(p + 12))), true);    // ELLIPSE
            case 44 -> shape(new RoundRectangle2D.Double(Math.min(i32(p + 8), i32(p + 16)), Math.min(i32(p + 12), i32(p + 20)),
                    Math.abs(i32(p + 16) - i32(p + 8)), Math.abs(i32(p + 20) - i32(p + 12)), i32(p + 24), i32(p + 28)), true); // ROUNDRECT
            case 45, 46, 47 -> arc(type, p);                                  // ARC, CHORD, PIE
            case 59 -> { path = new Path2D.Double(st.polyFill == 2 ? Path2D.WIND_NON_ZERO : Path2D.WIND_EVEN_ODD); Point2D q = pt(st.cpx, st.cpy); path.moveTo(q.getX(), q.getY()); } // BEGINPATH
            case 60 -> { lastPath = path; path = null; }                      // ENDPATH
            case 61 -> { if (path != null) path.closePath(); }               // CLOSEFIGURE
            case 62 -> { if (lastPath != null) fill(lastPath); lastPath = null; }                       // FILLPATH
            case 63 -> { if (lastPath != null) { fill(lastPath); stroke(lastPath); } lastPath = null; } // STROKEANDFILLPATH
            case 64 -> { if (lastPath != null) stroke(lastPath); lastPath = null; }                     // STROKEPATH
            case 67 -> { if (lastPath != null) setClip(new Area(lastPath), i32(p + 8)); lastPath = null; } // SELECTCLIPPATH
            case 30, 29 -> {                                                  // INTERSECTCLIPRECT, EXCLUDECLIPRECT
                Shape r = logToOut().createTransformedShape(new Rectangle2D.Double(i32(p + 8), i32(p + 12),
                        i32(p + 16) - i32(p + 8), i32(p + 20) - i32(p + 12)));
                setClip(new Area(r), type == 30 ? 1 : 4);
            }
            case 75 -> clipRegion(p);                                         // EXTSELECTCLIPRGN
            case 83, 84 -> text(p, type == 84);                               // EXTTEXTOUTA/W
            case 76 -> bitBlt(p, false);                                      // BITBLT
            case 77 -> bitBlt(p, true);                                       // STRETCHBLT
            case 81 -> stretchDib(p);                                         // STRETCHDIBITS
            case 80 -> setDibToDevice(p);                                     // SETDIBITSTODEVICE
            default -> { }
        }
    }

    private AffineTransform xform(int p) {
        return new AffineTransform(f32(p), f32(p + 4), f32(p + 8), f32(p + 12), f32(p + 16), f32(p + 20));
    }

    private LogFont logFont(int p) {
        int h = i32(p), esc = i32(p + 8), weight = i32(p + 16);
        boolean it = b.get(p + 20) != 0, ul = b.get(p + 21) != 0, so = b.get(p + 22) != 0;
        byte[] fn = new byte[64];
        b.get(p + 28, fn);
        String face = new String(fn, StandardCharsets.UTF_16LE);
        int z = face.indexOf('\0');
        if (z >= 0) face = face.substring(0, z);
        return new LogFont(h, esc, weight, it, ul, so, face);
    }

    private void select(int ih) {
        if ((ih & 0x80000000) != 0) {
            // 標準オブジェクト
            switch (ih & 0x7FFFFFFF) {
                case 0 -> st.brush = new Brush(0, Color.WHITE);
                case 1 -> st.brush = new Brush(0, new Color(192, 192, 192));
                case 2 -> st.brush = new Brush(0, new Color(128, 128, 128));
                case 3 -> st.brush = new Brush(0, new Color(64, 64, 64));
                case 4 -> st.brush = new Brush(0, Color.BLACK);
                case 5 -> st.brush = new Brush(1, Color.WHITE);
                case 6 -> st.pen = new Pen(0, 0, Color.WHITE);
                case 7 -> st.pen = new Pen(0, 0, Color.BLACK);
                case 8 -> st.pen = new Pen(5, 0, Color.BLACK);
                default -> { }
            }
            return;
        }
        Object o = objects.get(ih);
        if (o instanceof Pen pn) st.pen = pn;
        else if (o instanceof Brush br) st.brush = br;
        else if (o instanceof LogFont lf) st.font = lf;
    }

    // ---------------------------------------------------------------- 図形

    private void lineTo(double x, double y) {
        Point2D a = pt(st.cpx, st.cpy), c = pt(x, y);
        if (path != null) path.lineTo(c.getX(), c.getY());
        else stroke(new Line2D.Double(a, c));
        st.cpx = x; st.cpy = y;
    }

    /** POLY* の記録 (kind: 2 ベジェ, 3 多角形, 4 折れ線, 5 ベジェ (現在位置から), 6 折れ線 (現在位置から))。 */
    private void poly(int kind, int p, boolean small) {
        int n = i32(p + 24);
        int q = p + 28;
        Path2D sh = new Path2D.Double(st.polyFill == 2 ? Path2D.WIND_NON_ZERO : Path2D.WIND_EVEN_ODD);
        boolean fromCp = kind == 5 || kind == 6;
        Path2D target = path != null ? path : sh;
        if (fromCp) {
            if (path == null) { Point2D c = pt(st.cpx, st.cpy); sh.moveTo(c.getX(), c.getY()); }
        }
        double lx = st.cpx, ly = st.cpy;
        for (int k = 0; k < n; ) {
            double x = small ? i16(q + k * 4) : i32(q + k * 8), y = small ? i16(q + k * 4 + 2) : i32(q + k * 8 + 4);
            boolean bez = kind == 2 || kind == 5;
            if (k == 0 && !fromCp) {
                Point2D c = pt(x, y);
                target.moveTo(c.getX(), c.getY());
                k++;
                lx = x; ly = y;
                continue;
            }
            if (bez && k + 2 < n + (fromCp ? 0 : 0)) {
                double x2 = small ? i16(q + (k + 1) * 4) : i32(q + (k + 1) * 8), y2 = small ? i16(q + (k + 1) * 4 + 2) : i32(q + (k + 1) * 8 + 4);
                double x3 = small ? i16(q + (k + 2) * 4) : i32(q + (k + 2) * 8), y3 = small ? i16(q + (k + 2) * 4 + 2) : i32(q + (k + 2) * 8 + 4);
                Point2D c1 = pt(x, y), c2 = pt(x2, y2), c3 = pt(x3, y3);
                target.curveTo(c1.getX(), c1.getY(), c2.getX(), c2.getY(), c3.getX(), c3.getY());
                lx = x3; ly = y3;
                k += 3;
            } else {
                Point2D c = pt(x, y);
                target.lineTo(c.getX(), c.getY());
                lx = x; ly = y;
                k++;
            }
        }
        if (fromCp) { st.cpx = lx; st.cpy = ly; }
        if (path != null) return;
        if (kind == 3) {
            sh.closePath();
            fill(sh);
        }
        stroke(sh);
        drawn++;
    }

    private void polyPoly(boolean polygon, int p, boolean small) {
        int np = i32(p + 24);
        int q = p + 32;
        if (np <= 0 || (long) np * 4 > i32(p + 4)) return;  // 記録の大きさを超える数は壊れた記録
        int[] counts = new int[np];
        for (int k = 0; k < np; k++) counts[k] = i32(q + k * 4);
        q += np * 4;
        Path2D sh = new Path2D.Double(st.polyFill == 2 ? Path2D.WIND_NON_ZERO : Path2D.WIND_EVEN_ODD);
        int idx = 0;
        for (int c : counts) {
            for (int k = 0; k < c; k++, idx++) {
                double x = small ? i16(q + idx * 4) : i32(q + idx * 8), y = small ? i16(q + idx * 4 + 2) : i32(q + idx * 8 + 4);
                Point2D o = pt(x, y);
                if (k == 0) sh.moveTo(o.getX(), o.getY());
                else sh.lineTo(o.getX(), o.getY());
            }
            if (polygon) sh.closePath();
        }
        if (path != null) { path.append(sh, false); return; }
        if (polygon) fill(sh);
        stroke(sh);
        drawn++;
    }

    private void arc(int type, int p) {
        double l = i32(p + 8), t = i32(p + 12), r = i32(p + 16), bt = i32(p + 20);
        double cx = (l + r) / 2, cy = (t + bt) / 2;
        double a1 = Math.toDegrees(Math.atan2(-(i32(p + 28) - cy), i32(p + 24) - cx));
        double a2 = Math.toDegrees(Math.atan2(-(i32(p + 36) - cy), i32(p + 32) - cx));
        double ext = a2 - a1;
        if (ext <= 0) ext += 360;
        int kind = type == 45 ? Arc2D.OPEN : type == 46 ? Arc2D.CHORD : Arc2D.PIE;
        // GDI の円弧は反時計回り (y 下向きの座標で)
        Arc2D arc = new Arc2D.Double(Math.min(l, r), Math.min(t, bt), Math.abs(r - l), Math.abs(bt - t), a1, ext, kind);
        shape(arc, type != 45);
    }

    /** 論理座標の図形を、塗り (fillIt なら) と線で描く。 */
    private void shape(Shape s, boolean fillIt) {
        Shape o = logToOut().createTransformedShape(s);
        if (path != null) { path.append(o, false); return; }
        if (fillIt) fill(o);
        stroke(o);
        drawn++;
    }

    private void fill(Shape o) {
        if (st.brush.style == 1) return;  // BS_NULL
        if (st.brush.paint != null) g.setPaint(st.brush.paint); else g.setColor(st.brush.color);
        g.fill(o);
        drawn++;
    }

    /**
     * 模様のブラシ (CREATEDIBPATTERNBRUSHPT・CREATEMONOBRUSH): DIB を出力の左上から 1 デバイス画素 = devPx で敷き詰める。
     * 色は画像の平均 (模様を描けないところ用)。
     */
    private void patternBrush(int p) {
        int ih = i32(p + 8), offBmi = i32(p + 16), cbBmi = i32(p + 20), offBits = i32(p + 24), cbBits = i32(p + 28);
        if (cbBmi <= 0 || offBmi <= 0 || offBits <= 0) return;
        BufferedImage img = dib(p + offBmi, cbBmi, p + offBits, cbBits, 0x00CC0020);
        if (img == null) return;
        long r = 0, gg = 0, bb = 0;
        int n = img.getWidth() * img.getHeight();
        for (int y = 0; y < img.getHeight(); y++) for (int x = 0; x < img.getWidth(); x++) {
            int v = img.getRGB(x, y);
            r += v >> 16 & 255; gg += v >> 8 & 255; bb += v & 255;
        }
        Color avg = new Color((int) (r / n), (int) (gg / n), (int) (bb / n));
        java.awt.TexturePaint tp = new java.awt.TexturePaint(img, new Rectangle2D.Double(0, 0, img.getWidth() * devPx, img.getHeight() * devPx));
        objects.put(ih, new Brush(3, avg, tp));
    }

    /**
     * GRADIENTFILL: 頂点 (TRIVERTEX: x, y と 16 ビットの赤・緑・青) と、矩形 (横 0・縦 1) か三角形 (2) の組。
     * 矩形は 2 色の線形のグラデーション、三角形は細かい三角形に分けて頂点の色を補間する。
     */
    private void gradientFill(int p) {
        int nVer = i32(p + 24), nTri = i32(p + 28), mode = i32(p + 32);
        int recSize = i32(p + 4);
        int v0 = p + 36;
        int m0 = v0 + nVer * 16;
        int per = mode == 2 ? 12 : 8;
        if (nVer <= 0 || nTri <= 0 || (long) nVer * 16 + (long) nTri * per + 36 > recSize) return;
        Point2D[] pts = new Point2D[nVer];
        Color[] cols = new Color[nVer];
        for (int k = 0; k < nVer; k++) {
            int q = v0 + k * 16;
            pts[k] = pt(i32(q), i32(q + 4));
            cols[k] = new Color((i16(q + 8) & 0xFFFF) >> 8, (i16(q + 10) & 0xFFFF) >> 8, (i16(q + 12) & 0xFFFF) >> 8);
        }
        for (int t = 0; t < nTri; t++) {
            int q = m0 + t * per;
            if (mode == 2) {
                int a = i32(q), b2 = i32(q + 4), c = i32(q + 8);
                if (a < 0 || b2 < 0 || c < 0 || a >= nVer || b2 >= nVer || c >= nVer) continue;
                gouraud(pts[a], pts[b2], pts[c], cols[a], cols[b2], cols[c]);
            } else {
                int a = i32(q), b2 = i32(q + 4);
                if (a < 0 || b2 < 0 || a >= nVer || b2 >= nVer) continue;
                Point2D ul = pts[a], lr = pts[b2];
                Rectangle2D rc = new Rectangle2D.Double(Math.min(ul.getX(), lr.getX()), Math.min(ul.getY(), lr.getY()),
                        Math.abs(lr.getX() - ul.getX()), Math.abs(lr.getY() - ul.getY()));
                java.awt.GradientPaint gp = mode == 0
                        ? new java.awt.GradientPaint((float) ul.getX(), 0, cols[a], (float) lr.getX(), 0, cols[b2])
                        : new java.awt.GradientPaint(0, (float) ul.getY(), cols[a], 0, (float) lr.getY(), cols[b2]);
                g.setPaint(gp);
                g.fill(rc);
            }
            drawn++;
        }
    }

    /** 三角形を 16 × 16 に分けて、頂点の色を補間して塗る。 */
    private void gouraud(Point2D a, Point2D b2, Point2D c, Color ca, Color cb, Color cc) {
        int n = 16;
        for (int i = 0; i < n; i++) {
            for (int j = 0; j < n - i; j++) {
                for (int up = 0; up < (j < n - i - 1 ? 2 : 1); up++) {
                    double[][] w = up == 0
                            ? new double[][] {{i, j}, {i + 1, j}, {i, j + 1}}
                            : new double[][] {{i + 1, j}, {i + 1, j + 1}, {i, j + 1}};
                    Path2D tri = new Path2D.Double();
                    double sr = 0, sg = 0, sb = 0;
                    for (int k = 0; k < 3; k++) {
                        double u = w[k][0] / n, v = w[k][1] / n, x = 1 - u - v;
                        double px = x * a.getX() + u * b2.getX() + v * c.getX(), py = x * a.getY() + u * b2.getY() + v * c.getY();
                        if (k == 0) tri.moveTo(px, py); else tri.lineTo(px, py);
                        sr += x * ca.getRed() + u * cb.getRed() + v * cc.getRed();
                        sg += x * ca.getGreen() + u * cb.getGreen() + v * cc.getGreen();
                        sb += x * ca.getBlue() + u * cb.getBlue() + v * cc.getBlue();
                    }
                    tri.closePath();
                    g.setColor(new Color(clamp(sr / 3), clamp(sg / 3), clamp(sb / 3)));
                    g.fill(tri);
                    g.draw(tri);   // 隣との隙間を埋める
                }
            }
        }
    }

    private static int clamp(double v) { return (int) Math.max(0, Math.min(255, Math.round(v))); }

    private void stroke(Shape o) {
        int style = st.pen.style & 0xF;
        if (style == 5) return;  // PS_NULL
        double w = st.pen.width <= 0 ? devPx : Math.max(devPx, st.pen.width * logScale());
        int cap = switch (st.pen.style & 0xF00) { case 0x100 -> BasicStroke.CAP_SQUARE; case 0x200 -> BasicStroke.CAP_BUTT; default -> BasicStroke.CAP_ROUND; };
        int join = switch (st.pen.style & 0xF000) { case 0x1000 -> BasicStroke.JOIN_BEVEL; case 0x2000 -> BasicStroke.JOIN_MITER; default -> BasicStroke.JOIN_ROUND; };
        if ((st.pen.style & 0x10000) == 0) { cap = BasicStroke.CAP_ROUND; join = BasicStroke.JOIN_ROUND; } // コスメティックペン
        float[] dash = switch (style) {
            case 1 -> new float[] {(float) (6 * w), (float) (2 * w)};
            case 2 -> new float[] {(float) w, (float) w};
            case 3 -> new float[] {(float) (6 * w), (float) (2 * w), (float) w, (float) (2 * w)};
            case 4 -> new float[] {(float) (6 * w), (float) (2 * w), (float) w, (float) (2 * w), (float) w, (float) (2 * w)};
            default -> null;
        };
        g.setColor(st.pen.color);
        g.setStroke(new BasicStroke((float) w, dash != null ? BasicStroke.CAP_BUTT : cap, join, 10f, dash, 0));
        g.draw(o);
        drawn++;
    }

    // ---------------------------------------------------------------- クリップ

    /** クリップ領域を結合する (mode: 1 AND, 2 OR, 3 XOR, 4 DIFF, 5 COPY)。 */
    private void setClip(Area a, int mode) {
        Area cur = st.clip;
        switch (mode) {
            case 1 -> { if (cur == null) cur = a; else cur.intersect(a); }
            case 2 -> { if (cur != null) cur.add(a); }
            case 3 -> { if (cur != null) cur.exclusiveOr(a); else cur = a; }
            case 4 -> {
                if (cur == null) { cur = new Area(new Rectangle2D.Double(-1e6, -1e6, 2e6, 2e6)); }
                cur.subtract(a);
            }
            default -> cur = a;
        }
        st.clip = cur;
        applyClip();
    }

    private void applyClip() {
        g.setClip(st.clip);
    }

    private void clipRegion(int p) {
        int cb = i32(p + 8), mode = i32(p + 12);
        if (cb == 0) {
            if (mode == 5) { st.clip = null; applyClip(); }
            return;
        }
        // RGNDATA: ヘッダー 32 バイト (nCount はオフセット 8) の後に矩形 (デバイス単位)
        int q = p + 16, n = i32(q + 8);
        Area a = new Area();
        for (int k = 0; k < n; k++) {
            int r = q + 32 + k * 16;
            Rectangle2D rc = new Rectangle2D.Double(i32(r), i32(r + 4), i32(r + 8) - i32(r), i32(r + 12) - i32(r + 4));
            a.add(new Area(devToOut.createTransformedShape(rc)));
        }
        setClip(a, mode);
    }

    // ---------------------------------------------------------------- 文字

    private void text(int p, boolean wide) {
        int ref = p + 36;  // EMRTEXT
        double rx = i32(ref), ry = i32(ref + 4);
        int n = i32(ref + 8), offStr = i32(ref + 12), opts = i32(ref + 16), offDx = i32(ref + 36);
        // 文字数は記録の大きさの範囲内でなければならない (壊れた・悪意のある記録で巨大な配列を確保しない)
        int recSize = i32(p + 4);
        if (n <= 0 || offStr < 0 || (long) offStr + (long) n * (wide ? 2 : 1) > recSize) return;
        String s;
        if (wide) {
            byte[] bs = new byte[n * 2];
            b.get(p + offStr, bs);
            s = new String(bs, StandardCharsets.UTF_16LE);
        } else {
            byte[] bs = new byte[n];
            b.get(p + offStr, bs);
            s = new String(bs, java.nio.charset.Charset.forName("MS932"));
        }
        if ((opts & 0x2) != 0) {
            // ETO_OPAQUE: 矩形を背景色で塗る
            Shape r = logToOut().createTransformedShape(new Rectangle2D.Double(i32(ref + 20), i32(ref + 24),
                    i32(ref + 28) - i32(ref + 20), i32(ref + 32) - i32(ref + 24)));
            g.setColor(st.bkColor);
            g.fill(r);
        }
        boolean glyphs = (opts & 0x10) != 0 && wide;   // ETO_GLYPH_INDEX: 16 ビットの字形番号 (選んだフォントの番号)
        LogFont lf = st.font;
        String face = lf.face.startsWith("@") ? lf.face.substring(1) : lf.face;
        Font base = fonts.awtFont(face, lf.weight >= 600, lf.italic);
        // 文字の高さ: 負なら文字の高さ (em)、正ならセルの高さ (アセント + ディセント)
        java.awt.font.FontRenderContext frc = g.getFontRenderContext();
        java.awt.font.LineMetrics lm1 = base.deriveFont(100f).getLineMetrics("あ", frc);
        double cellPer = (lm1.getAscent() + lm1.getDescent()) / 100.0;
        double em = lf.height < 0 ? -lf.height : lf.height == 0 ? 12 : lf.height / cellPer;
        AffineTransform m = logToOut();
        double sy = Math.hypot(m.getShearX(), m.getScaleY()), sx = Math.hypot(m.getScaleX(), m.getShearY());
        double sizeOut = em * sy;
        if (sizeOut < 0.5) return;
        double esc = Math.toRadians(lf.escapement / 10.0);
        // 送り幅 (論理単位): Dx がなければフォントの字送り
        int[] dx = new int[s.length()];
        int dxStep = (opts & 0x2000) != 0 ? 2 : 1;  // ETO_PDY: x と y の組
        boolean haveDx = offDx > 0 && offDx + 4L * n * dxStep <= (long) i32(p + 4);
        double total = 0;
        Font fOut = base.deriveFont((float) sizeOut);
        for (int k = 0; k < s.length(); k++) {
            if (haveDx && k < n) dx[k] = i32(p + offDx + 4 * dxStep * Math.min(k, n - 1));
            else if (glyphs) dx[k] = (int) Math.round(fOut.createGlyphVector(frc, new int[] {s.charAt(k)}).getGlyphMetrics(0).getAdvance() / Math.max(1e-6, sx));
            else dx[k] = (int) Math.round(fOut.getStringBounds(s.substring(k, k + 1), frc).getWidth() / Math.max(1e-6, sx));
            total += dx[k];
        }
        boolean updateCp = (st.textAlign & 1) != 0;
        if (updateCp) { rx = st.cpx; ry = st.cpy; }
        int ha = st.textAlign & 6, va = st.textAlign & 24;
        double startX = ha == 2 ? -total : ha == 6 ? -total / 2 : 0;
        java.awt.font.LineMetrics lm = fOut.getLineMetrics("あ", frc);
        double baseOff = va == 24 ? 0 : va == 8 ? -lm.getDescent() : lm.getAscent();
        Point2D o = m.transform(new Point2D.Double(rx, ry), null);
        // 縦書きのフォント (@ で始まる) は文字を 90 度回す
        boolean vertical = lf.face.startsWith("@");
        double cos = Math.cos(-esc), sin = Math.sin(-esc);
        g.setColor(st.textColor);
        double adv = startX;
        for (int k = 0; k < s.length(); k++) {
            String ch = s.substring(k, k + 1);
            double ox = o.getX() + (adv * sx) * cos - baseOff * sin, oy = o.getY() + (adv * sx) * sin + baseOff * cos;
            AffineTransform save = g.getTransform();
            g.translate(ox, oy);
            g.rotate(-esc + (vertical ? Math.PI / 2 : 0));
            g.setFont(fOut);
            if (glyphs) g.drawGlyphVector(fOut.createGlyphVector(frc, new int[] {s.charAt(k)}), 0f, 0f);
            else g.drawString(ch, 0f, 0f);
            g.setTransform(save);
            adv += dx[k];
        }
        if (lf.underline || lf.strike) {
            double y = lf.underline ? lm.getUnderlineOffset() : lm.getStrikethroughOffset();
            double t = Math.max(devPx, lm.getUnderlineThickness());
            AffineTransform save = g.getTransform();
            g.translate(o.getX(), o.getY());
            g.rotate(-esc);
            g.setStroke(new BasicStroke((float) t));
            g.draw(new Line2D.Double(startX * sx, baseOff + y, (startX + total) * sx, baseOff + y));
            g.setTransform(save);
        }
        if (updateCp) st.cpx = rx + total;
        drawn++;
    }

    // ---------------------------------------------------------------- 画像

    private void bitBlt(int p, boolean stretch) {
        double xd = i32(p + 24), yd = i32(p + 28), cxd = i32(p + 32), cyd = i32(p + 36);
        int rop = i32(p + 40);
        // xformSrc (24 バイト) と crBkColorSrc, iUsageSrc の後に、BMI と画素の位置
        int offBmi = i32(p + 84), cbBmi = i32(p + 88), offBits = i32(p + 92), cbBits = i32(p + 96);
        int cxs = stretch ? i32(p + 100) : (int) cxd, cys = stretch ? i32(p + 104) : (int) cyd;
        Shape dest = logToOut().createTransformedShape(new Rectangle2D.Double(xd, yd, cxd, cyd));
        if (cbBmi == 0) {
            // 画像のない BITBLT: ブラシ (PATCOPY) や白・黒で塗る
            Color c = switch (rop) {
                case 0x00F00021 -> st.brush.style == 1 ? null : st.brush.color;
                case 0x00FF0062 -> Color.WHITE;
                case 0x00000042 -> Color.BLACK;
                default -> null;
            };
            if (c != null) {
                if (rop == 0x00F00021 && st.brush.paint != null) g.setPaint(st.brush.paint); else g.setColor(c);   // 模様のブラシは模様で
                g.fill(dest);
                drawn++;
            }
            return;
        }
        BufferedImage img = dib(p + offBmi, cbBmi, p + offBits, cbBits, rop);
        if (img == null) return;
        int xs = i32(p + 44), ys = i32(p + 48);
        drawImage(img, xs, ys, stretch ? cxs : img.getWidth(), stretch ? cys : img.getHeight(), xd, yd, cxd, cyd);
    }

    private void stretchDib(int p) {
        double xd = i32(p + 24), yd = i32(p + 28);
        int xs = i32(p + 32), ys = i32(p + 36), cxs = i32(p + 40), cys = i32(p + 44);
        int offBmi = i32(p + 48), cbBmi = i32(p + 52), offBits = i32(p + 56), cbBits = i32(p + 60);
        int rop = i32(p + 68);
        double cxd = i32(p + 72), cyd = i32(p + 76);
        BufferedImage img = dib(p + offBmi, cbBmi, p + offBits, cbBits, rop);
        if (img == null) return;
        // 下から上の DIB では、元の範囲の y は下端から数える
        if (i32(p + offBmi + 8) > 0) ys = img.getHeight() - ys - cys;
        drawImage(img, xs, ys, cxs, cys, xd, yd, cxd, cyd);
    }

    private void setDibToDevice(int p) {
        double xd = i32(p + 24), yd = i32(p + 28);
        int xs = i32(p + 32), ys = i32(p + 36), cxs = i32(p + 40), cys = i32(p + 44);
        int offBmi = i32(p + 48), cbBmi = i32(p + 52), offBits = i32(p + 56), cbBits = i32(p + 60);
        BufferedImage img = dib(p + offBmi, cbBmi, p + offBits, cbBits, 0x00CC0020);
        if (img == null) return;
        if (i32(p + offBmi + 8) > 0) ys = img.getHeight() - ys - cys;
        drawImage(img, xs, ys, cxs, cys, xd, yd, cxs, cys);
    }

    /** 画像の一部 (xs, ys, cxs, cys) を論理座標の矩形 (xd, yd, cxd, cyd) に描く。 */
    private void drawImage(BufferedImage img, int xs, int ys, int cxs, int cys, double xd, double yd, double cxd, double cyd) {
        if (cxs <= 0 || cys <= 0) return;
        xs = Math.max(0, Math.min(xs, img.getWidth() - 1));
        ys = Math.max(0, Math.min(ys, img.getHeight() - 1));
        cxs = Math.min(cxs, img.getWidth() - xs);
        cys = Math.min(cys, img.getHeight() - ys);
        if (cxs <= 0 || cys <= 0) return;
        BufferedImage sub = img.getSubimage(xs, ys, cxs, cys);
        AffineTransform t = logToOut();
        t.translate(xd, yd);
        t.scale(cxd / cxs, cyd / cys);
        g.drawImage(sub, t, null);
        drawn++;
    }

    /**
     * DIB (BITMAPINFOHEADER + 画素) を画像にする。rop が SRCAND なら白を、SRCPAINT なら黒を透明にする (重ね合わせの近似)。
     */
    private BufferedImage dib(int bmi, int cbBmi, int bits, int cbBits, int rop) {
        int hs = i32(bmi);
        int w = i32(bmi + 4), h = i32(bmi + 8);
        int bpp = i16(bmi + 14) & 0xFFFF, comp = i32(bmi + 16);
        if (w <= 0 || h == 0 || !Images.pixelsAllowed(w, Math.abs((long) h))) return null;
        if (comp == 4 || comp == 5) {
            // JPEG / PNG
            try {
                if (cbBits <= 0 || bits + (long) cbBits > b.capacity()) return null;
                byte[] d = new byte[cbBits];
                b.get(bits, d);
                return Images.decodeLimited(d);
            } catch (Exception e) {
                return null;
            }
        }
        boolean topDown = h < 0;
        h = Math.abs(h);
        if ((comp == 1 && bpp == 8) || (comp == 2 && bpp == 4)) return rle(bmi, hs, bpp, w, h, bits, cbBits, i32(bmi + 32));
        int clrUsed = i32(bmi + 32);
        // パレットは 1 << bpp 色まで (それを超える色数や負の値は壊れた記録)
        int nPal = bpp <= 8 ? (clrUsed > 0 && clrUsed <= 1 << bpp ? clrUsed : 1 << bpp) : 0;
        int palOff = bmi + hs + (comp == 3 && hs == 40 ? 12 : 0);
        int[] pal = new int[Math.max(nPal, 1)];
        for (int k = 0; k < nPal && palOff + k * 4 + 3 < b.capacity(); k++) {
            int v = i32(palOff + k * 4);
            pal[k] = 0xFF000000 | ((v >> 16 & 255) << 16) | ((v >> 8 & 255) << 8) | (v & 255);
        }
        int stride = ((w * bpp + 31) / 32) * 4;
        if (bits + (long) stride * h > b.capacity()) return null;
        BufferedImage img = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        for (int y = 0; y < h; y++) {
            int row = bits + (topDown ? y : h - 1 - y) * stride;
            for (int x = 0; x < w; x++) {
                int argb;
                switch (bpp) {
                    case 1 -> argb = pal[Math.min(pal.length - 1, (b.get(row + x / 8) >> (7 - x % 8)) & 1)];
                    case 4 -> argb = pal[Math.min(pal.length - 1, (b.get(row + x / 2) >> (x % 2 == 0 ? 4 : 0)) & 15)];
                    case 8 -> argb = pal[Math.min(pal.length - 1, b.get(row + x) & 255)];
                    case 16 -> {
                        int v = i16(row + x * 2) & 0xFFFF;
                        argb = 0xFF000000 | (((v >> 10) & 31) * 255 / 31) << 16 | (((v >> 5) & 31) * 255 / 31) << 8 | ((v & 31) * 255 / 31);
                    }
                    case 24 -> argb = 0xFF000000 | (b.get(row + x * 3 + 2) & 255) << 16 | (b.get(row + x * 3 + 1) & 255) << 8 | (b.get(row + x * 3) & 255);
                    case 32 -> argb = 0xFF000000 | (b.get(row + x * 4 + 2) & 255) << 16 | (b.get(row + x * 4 + 1) & 255) << 8 | (b.get(row + x * 4) & 255);
                    default -> { return null; }
                }
                int rgb = argb & 0xFFFFFF;
                if (rop == 0x008800C6 && rgb == 0xFFFFFF) argb = 0;      // SRCAND: 白は下を残す
                else if (rop == 0x00EE0086 && rgb == 0) argb = 0;         // SRCPAINT: 黒は下を残す
                else if (rop == 0x00330008) argb = 0xFF000000 | (~rgb & 0xFFFFFF); // NOTSRCCOPY
                img.setRGB(x, y, argb);
            }
        }
        return img;
    }

    /**
     * RLE で圧縮した DIB (BI_RLE8・BI_RLE4)。下の行から上へ。符号化されない画素は透明 (下を残す)。
     */
    private BufferedImage rle(int bmi, int hs, int bpp, int w, int h, int bits, int cbBits, int clrUsed) {
        int nPal = clrUsed > 0 && clrUsed <= 1 << bpp ? clrUsed : 1 << bpp;
        int[] pal = new int[nPal];
        for (int k = 0; k < nPal && bmi + hs + k * 4 + 3 < b.capacity(); k++) {
            int v = i32(bmi + hs + k * 4);
            pal[k] = 0xFF000000 | (v & 0xFFFFFF);
        }
        if (cbBits <= 0 || bits + (long) cbBits > b.capacity()) return null;
        BufferedImage img = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        int x = 0, y = 0, q = bits, end = bits + cbBits;
        while (q + 1 < end && y < h) {
            int n = b.get(q) & 255, c = b.get(q + 1) & 255;
            q += 2;
            if (n > 0) {
                for (int k = 0; k < n; k++, x++) {
                    int idx = bpp == 8 ? c : (k % 2 == 0 ? c >> 4 : c & 15);
                    if (x < w) img.setRGB(x, h - 1 - y, pal[Math.min(nPal - 1, idx)]);
                }
            } else if (c == 0) { x = 0; y++; }
            else if (c == 1) break;
            else if (c == 2) { if (q + 1 >= end) break; x += b.get(q) & 255; y += b.get(q + 1) & 255; q += 2; }
            else {
                // 絶対モード: c 画素をそのまま (2 バイト境界に揃える)
                int bytes = bpp == 8 ? c : (c + 1) / 2;
                if (q + bytes > end) break;
                for (int k = 0; k < c; k++, x++) {
                    int v = bpp == 8 ? b.get(q + k) & 255 : (k % 2 == 0 ? (b.get(q + k / 2) >> 4) & 15 : b.get(q + k / 2) & 15);
                    if (x < w) img.setRGB(x, h - 1 - y, pal[Math.min(nPal - 1, v)]);
                }
                q += bytes + (bytes & 1);
            }
        }
        return img;
    }
}
