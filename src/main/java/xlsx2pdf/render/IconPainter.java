package xlsx2pdf.render;

import java.awt.Color;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Arc2D;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.awt.geom.RoundRectangle2D;
import java.io.IOException;
import java.util.Map;

/**
 * 条件付き書式のアイコンを図形で描く。形と色は Excel の印刷に合わせて作ったもの (Excel の画像は使わない)。
 * 座標は 24 × 24 の格子で定義し、描く大きさに拡大する。
 */
final class IconPainter {
    private IconPainter() {}

    // 塗りと縁の色
    private static final Color RED = new Color(0xD65532), RED_LN = new Color(0x9E371C);
    private static final Color YEL = new Color(0xEAC282), YEL_LN = new Color(0xA4802B);
    private static final Color GRN = new Color(0x68A490), GRN_LN = new Color(0x326F5B);
    private static final Color GRY = new Color(0x808080), GRY_LN = new Color(0x646262);
    private static final Color DRK = new Color(0x505050), DRK_LN = new Color(0x303030);
    private static final Color LGR = new Color(0xB3B3B3), LGR_LN = new Color(0x808080);
    private static final Color PNK = new Color(0xEDB2A3), PNK_LN = new Color(0xD86344);
    private static final Color BLU = new Color(0x4A7DB1), BLU_LN = new Color(0x335F8C);

    /** セットごとのアイコン (低い値から順)。 */
    private static final Map<String, String[]> SETS = Map.ofEntries(
            Map.entry("3Arrows", new String[] {"aD:r", "aR:y", "aU:g"}),
            Map.entry("3ArrowsGray", new String[] {"aD:k", "aR:k", "aU:k"}),
            Map.entry("3Flags", new String[] {"flag:r", "flag:y", "flag:g"}),
            Map.entry("3TrafficLights1", new String[] {"dot:r", "dot:y", "dot:g"}),
            Map.entry("3TrafficLights2", new String[] {"rim:r", "rim:y", "rim:g"}),
            Map.entry("3Signs", new String[] {"diamond:r", "tri:y", "dot:g"}),
            Map.entry("3Symbols", new String[] {"cX:r", "cEx:y", "cCk:g"}),
            Map.entry("3Symbols2", new String[] {"X:r", "Ex:y", "Ck:g"}),
            Map.entry("4Arrows", new String[] {"aD:r", "aDR:y", "aUR:y", "aU:g"}),
            Map.entry("4ArrowsGray", new String[] {"aD:k", "aDR:k", "aUR:k", "aU:k"}),
            Map.entry("4RedToBlack", new String[] {"dot:d", "dot:l", "dot:p", "dot:r"}),
            Map.entry("4Rating", new String[] {"bars1", "bars2", "bars3", "bars4"}),
            Map.entry("4TrafficLights", new String[] {"dot:d", "dot:r", "dot:y", "dot:g"}),
            Map.entry("5Arrows", new String[] {"aD:r", "aDR:y", "aR:y", "aUR:y", "aU:g"}),
            Map.entry("5ArrowsGray", new String[] {"aD:k", "aDR:k", "aR:k", "aUR:k", "aU:k"}),
            Map.entry("5Rating", new String[] {"bars0", "bars1", "bars2", "bars3", "bars4"}),
            Map.entry("5Quarters", new String[] {"q0", "q1", "q2", "q3", "q4"}),
            Map.entry("3Stars", new String[] {"star0", "star1", "star2"}),
            Map.entry("3Triangles", new String[] {"tD:r", "dash:y", "tU:g"}),
            Map.entry("5Boxes", new String[] {"box0", "box1", "box2", "box3", "box4"}));

    /** アイコンセットの名前を知っているか。 */
    static boolean known(String set) { return SETS.containsKey(set); }

    /**
     * アイコンを描く。
     *
     * @param x 左端
     * @param y 上端
     * @param s 大きさ (正方形の一辺、pt)
     */
    static void paint(PdfCanvas cv, String set, int index, double x, double y, double s) throws IOException {
        String[] icons = SETS.get(set);
        if (icons == null || index < 0 || index >= icons.length) return;
        String code = icons[index];
        String shape = code, col = "";
        int k = code.indexOf(':');
        if (k > 0) { shape = code.substring(0, k); col = code.substring(k + 1); }
        Color fill = switch (col) {
            case "r" -> RED; case "y" -> YEL; case "g" -> GRN; case "k" -> GRY; case "d" -> DRK; case "l" -> LGR; case "p" -> PNK;
            default -> BLU;
        };
        Color line = switch (col) {
            case "r" -> RED_LN; case "y" -> YEL_LN; case "g" -> GRN_LN; case "k" -> GRY_LN; case "d" -> DRK_LN; case "l" -> LGR_LN;
            case "p" -> PNK_LN; default -> BLU_LN;
        };
        AffineTransform at = new AffineTransform();
        at.translate(x, y);
        at.scale(s / 24, s / 24);
        float lw = (float) (s / 24);
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(lw);
        switch (shape) {
            case "aU" -> shape(cv, at, arrow(0), fill, line, st);
            case "aR" -> shape(cv, at, arrow(90), fill, line, st);
            case "aD" -> shape(cv, at, arrow(180), fill, line, st);
            case "aUR" -> shape(cv, at, arrow(45), fill, line, st);
            case "aDR" -> shape(cv, at, arrow(135), fill, line, st);
            case "flag" -> {
                shape(cv, at, new Rectangle2D.Double(3, 1, 2.5, 22), GRY, GRY_LN, st);
                shape(cv, at, poly(5.5, 2, 21, 8, 5.5, 14), fill, line, st);
            }
            case "dot" -> shape(cv, at, new Ellipse2D.Double(1.5, 1.5, 21, 21), fill, line, st);
            case "rim" -> {
                shape(cv, at, new RoundRectangle2D.Double(1, 1, 22, 22, 6, 6), DRK, DRK_LN, st);
                shape(cv, at, new Ellipse2D.Double(5.5, 5.5, 13, 13), fill, line, st);
            }
            case "diamond" -> shape(cv, at, poly(12, 1.5, 22.5, 12, 12, 22.5, 1.5, 12), fill, line, st);
            case "tri" -> shape(cv, at, poly(12, 2, 22.5, 21, 1.5, 21), fill, line, st);
            case "cX", "cEx", "cCk" -> {
                shape(cv, at, new Ellipse2D.Double(1.5, 1.5, 21, 21), fill, line, st);
                glyph(cv, at, shape.substring(1), Color.WHITE, 2.6 * lw);
            }
            case "X", "Ex", "Ck" -> glyph(cv, at, shape, fill, 4.2 * lw);
            case "tU" -> shape(cv, at, poly(12, 6, 22, 18, 2, 18), fill, line, st);
            case "tD" -> shape(cv, at, poly(2, 6, 22, 6, 12, 18), fill, line, st);
            case "dash" -> shape(cv, at, new Rectangle2D.Double(2, 9, 20, 6), fill, line, st);
            default -> {
                if (shape.startsWith("bars")) bars(cv, at, shape.charAt(4) - '0', st);
                else if (shape.startsWith("q")) quarters(cv, at, shape.charAt(1) - '0', st);
                else if (shape.startsWith("star")) star(cv, at, shape.charAt(4) - '0', st);
                else if (shape.startsWith("box")) boxes(cv, at, shape.charAt(3) - '0', st);
            }
        }
    }

    /** 塗りと縁を描く。 */
    private static void shape(PdfCanvas cv, AffineTransform at, Shape sh, Color fill, Color line, PdfCanvas.Stroke st) throws IOException {
        Shape t = at.createTransformedShape(sh);
        if (fill != null) cv.fill(t, fill);
        if (line != null) cv.stroke(t, st, line);
    }

    /** 上向きの矢印を deg 度 (時計回り) 回したもの。斜めは枠に収まるよう縮める。 */
    private static Shape arrow(double deg) {
        Path2D p = poly(12, 1, 22.5, 11.5, 16, 11.5, 16, 23, 8, 23, 8, 11.5, 1.5, 11.5);
        AffineTransform r = new AffineTransform();
        r.translate(12, 12);
        r.rotate(Math.toRadians(deg));
        if (deg % 90 != 0) r.scale(0.85, 0.85);
        r.translate(-12, -12);
        return r.createTransformedShape(p);
    }

    /** 記号 (X, Ex = 感嘆符, Ck = チェック) を太い線で描く。 */
    private static void glyph(PdfCanvas cv, AffineTransform at, String g, Color c, double w) throws IOException {
        PdfCanvas.Stroke st = new PdfCanvas.Stroke((float) w);
        st.cap = 1;
        st.join = 1;
        boolean small = c == Color.WHITE; // 丸の中に描く
        Path2D p = new Path2D.Double();
        switch (g) {
            case "X" -> {
                double a = small ? 7.5 : 4, b = 24 - a;
                p.moveTo(a, a); p.lineTo(b, b); p.moveTo(b, a); p.lineTo(a, b);
            }
            case "Ck" -> {
                if (small) { p.moveTo(6.5, 12.5); p.lineTo(10.5, 16.5); p.lineTo(17.5, 8); }
                else { p.moveTo(3.5, 12.5); p.lineTo(9.5, 18.5); p.lineTo(20.5, 5); }
            }
            default -> {
                if (small) { p.moveTo(12, 6); p.lineTo(12, 13.5); p.moveTo(12, 17.5); p.lineTo(12, 18); }
                else { p.moveTo(12, 3); p.lineTo(12, 14.5); p.moveTo(12, 19.5); p.lineTo(12, 20.5); }
            }
        }
        cv.stroke(at.createTransformedShape(p), st, c);
    }

    /** 棒の評価 (4 本、filled 本を青で塗る)。 */
    private static void bars(PdfCanvas cv, AffineTransform at, int filled, PdfCanvas.Stroke st) throws IOException {
        for (int i = 0; i < 4; i++) {
            double h = 7 + i * 4.5;
            Rectangle2D r = new Rectangle2D.Double(2.5 + i * 5, 22 - h, 4.5, h);
            boolean on = i < filled;
            shape(cv, at, r, on ? BLU : LGR, on ? BLU_LN : LGR_LN, st);
        }
    }

    /** 丸の 4 分の q を塗る。 */
    private static void quarters(PdfCanvas cv, AffineTransform at, int q, PdfCanvas.Stroke st) throws IOException {
        Ellipse2D e = new Ellipse2D.Double(1.5, 1.5, 21, 21);
        shape(cv, at, e, Color.WHITE, null, st);
        if (q >= 4) shape(cv, at, e, DRK, null, st);
        else if (q > 0) shape(cv, at, new Arc2D.Double(1.5, 1.5, 21, 21, 90, -90 * q, Arc2D.PIE), DRK, null, st);
        shape(cv, at, e, null, DRK_LN, st);
    }

    /** 星 (0 空、1 半分、2 全部)。 */
    private static void star(PdfCanvas cv, AffineTransform at, int level, PdfCanvas.Stroke st) throws IOException {
        Path2D p = new Path2D.Double();
        for (int i = 0; i < 10; i++) {
            double r = i % 2 == 0 ? 11 : 4.6;
            double a = Math.toRadians(-90 + i * 36);
            double px = 12 + r * Math.cos(a), py = 12.8 + r * Math.sin(a);
            if (i == 0) p.moveTo(px, py); else p.lineTo(px, py);
        }
        p.closePath();
        shape(cv, at, p, level == 2 ? YEL : Color.WHITE, null, st);
        if (level == 1) {
            cv.save();
            cv.clip(at.createTransformedShape(new Rectangle2D.Double(0, 0, 12, 24)));
            shape(cv, at, p, YEL, null, st);
            cv.restore();
        }
        shape(cv, at, p, null, level == 0 ? GRY : YEL_LN, st);
    }

    /** 2 × 2 の箱を n 個青で塗る (左下、右下、左上、右上の順)。 */
    private static void boxes(PdfCanvas cv, AffineTransform at, int n, PdfCanvas.Stroke st) throws IOException {
        double[][] pos = {{2, 12}, {12, 12}, {2, 2}, {12, 2}};
        for (int i = 0; i < 4; i++) {
            boolean on = i < n;
            shape(cv, at, new Rectangle2D.Double(pos[i][0], pos[i][1], 10, 10), on ? BLU : LGR, on ? BLU_LN : LGR_LN, st);
        }
    }

    /** 座標の並びから閉じた多角形を作る。 */
    private static Path2D poly(double... c) {
        Path2D p = new Path2D.Double();
        p.moveTo(c[0], c[1]);
        for (int i = 2; i < c.length; i += 2) p.lineTo(c[i], c[i + 1]);
        p.closePath();
        return p;
    }
}
