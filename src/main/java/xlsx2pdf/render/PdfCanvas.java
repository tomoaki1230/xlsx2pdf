package xlsx2pdf.render;

import java.awt.Color;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.PathIterator;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.PDExtendedGraphicsState;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;
import org.apache.pdfbox.util.Matrix;

/**
 * PDFBox のコンテンツストリームを包む薄い描画層。座標は左上原点・下向きの y 軸で、単位は pt。
 */
public final class PdfCanvas implements AutoCloseable {
    /** 出力先の PDF。 */
    public final PDDocument doc;
    /** フォントの検索・埋め込み。 */
    public final FontManager fonts;
    private final PDPageContentStream cs;
    /** 透明度の組み合わせ (塗り × 1000 + 線) ごとのグラフィックス状態。 */
    private final Map<Integer, PDExtendedGraphicsState> alphaStates = new HashMap<>();

    /** 線のスタイル。 */
    public static final class Stroke {
        /** 線の太さ (pt)。 */
        public float width = 1;
        /** 破線のパターン (pt。null なら実線)。 */
        public float[] dash;
        /** 線端 (0 平ら、1 丸、2 四角)。 */
        public int cap = 0;
        /** 角 (0 マイター、1 丸、2 面取り)。 */
        public int join = 0;

        /** 指定の太さの実線。 */
        public Stroke(float width) { this.width = width; }

        /** 破線のパターンを設定して自身を返す。 */
        public Stroke dash(float... d) { this.dash = d; return this; }
    }

    private final PDPage page;
    private final float pageH;

    /**
     * ページに描くキャンバスを作る (座標系を左上原点に変換する)。
     */
    public PdfCanvas(PDDocument doc, PDPage page, FontManager fonts) throws IOException {
        this.doc = doc;
        this.fonts = fonts;
        this.page = page;
        if (page.getResources() == null) page.setResources(new org.apache.pdfbox.pdmodel.PDResources());
        this.cs = new PDPageContentStream(doc, page);
        float h = page.getMediaBox().getHeight();
        this.pageH = h;
        cs.transform(new Matrix(1, 0, 0, -1, 0, h));
    }

    /** グラフィックス状態を保存する。 */
    public void save() throws IOException { cs.saveGraphicsState(); depth++; }

    /** save の入れ子の深さ。 */
    private int depth;

    /** いまの save の深さ (例外のときに {@link #restoreTo} で戻すために覚える)。 */
    public int depth() { return depth; }

    /**
     * save の深さを d まで戻す。描く途中で例外が起きて restore が飛ばされても、クリップや変換が後の描画に残らないようにする。
     */
    public void restoreTo(int d) {
        try {
            while (depth > d) restore();
        } catch (IOException e) {
            // 書けなければ戻せない (ページの書き出しも失敗する)
        }
    }

    /** 文字の左右反転を打ち消す (右から左のシートで、本体を左右反転して描くとき)。 */
    public boolean mirrorText;

    /** x を軸に左右反転する変換。 */
    public static AffineTransform flipAbout(double x) {
        AffineTransform at = AffineTransform.getTranslateInstance(2 * x, 0);
        at.scale(-1, 1);
        return at;
    }

    /** グラフィックス状態を戻す。 */
    public void restore() throws IOException {
        if (depth == 0) return;   // 対応する save のない restore は書かない (PDF の q・Q の数を合わせる)
        cs.restoreGraphicsState();
        depth--;
    }

    /** 座標変換を追加する。 */
    public void transform(AffineTransform at) throws IOException {
        cs.transform(new Matrix(at));
    }

    /** 平行移動を追加する。 */
    public void translate(double x, double y) throws IOException {
        cs.transform(Matrix.getTranslateInstance((float) x, (float) y));
    }

    /** 拡大縮小を追加する。 */
    public void scale(double sx, double sy) throws IOException {
        cs.transform(Matrix.getScaleInstance((float) sx, (float) sy));
    }

    /** 矩形でクリップする。 */
    public void clipRect(double x, double y, double w, double h) throws IOException {
        cs.addRect((float) x, (float) y, (float) w, (float) h);
        cs.clip();
    }

    /** 図形でクリップする。 */
    public void clip(Shape s) throws IOException {
        path(s);
        cs.clip();
    }

    /** 塗りと線の透明度を設定する (一度も透明度を使っていなければ不透明の設定は省く)。 */
    private void alpha(float fill, float stroke) throws IOException {
        int key = Math.round(fill * 255) * 1000 + Math.round(stroke * 255);
        if (key == 255 * 1000 + 255 && alphaStates.isEmpty()) return;
        PDExtendedGraphicsState gs = alphaStates.get(key);
        if (gs == null) {
            gs = new PDExtendedGraphicsState();
            gs.setNonStrokingAlphaConstant(fill);
            gs.setStrokingAlphaConstant(stroke);
            alphaStates.put(key, gs);
        }
        cs.setGraphicsStateParameters(gs);
    }

    private static float alphaOf(Color c) { return c.getAlpha() / 255f; }

    /** 矩形を塗りつぶす (null・完全透明は描かない)。 */
    public void fillRect(double x, double y, double w, double h, Color c) throws IOException {
        if (c == null || c.getAlpha() == 0) return;
        alpha(alphaOf(c), 1);
        cs.setNonStrokingColor(rgb(c));
        cs.addRect((float) x, (float) y, (float) w, (float) h);
        cs.fill();
    }

    /**
     * 矩形を、(x0, y0) の色 c0 から (x1, y1) の色 c1 へ直線で変わるグラデーションで塗る (PDF の軸方向のシェーディング)。
     */
    public void fillAxial(double x, double y, double w, double h, double x0, double y0, double x1, double y1, Color c0, Color c1)
            throws IOException {
        if (w <= 0 || h <= 0) return;
        org.apache.pdfbox.cos.COSDictionary fn = new org.apache.pdfbox.cos.COSDictionary();
        fn.setInt(org.apache.pdfbox.cos.COSName.FUNCTION_TYPE, 2);
        fn.setItem(org.apache.pdfbox.cos.COSName.DOMAIN, floats(0, 1));
        float[] a = rgb(c0).getRGBColorComponents(null), b = rgb(c1).getRGBColorComponents(null);
        fn.setItem(org.apache.pdfbox.cos.COSName.C0, floats(a[0], a[1], a[2]));
        fn.setItem(org.apache.pdfbox.cos.COSName.C1, floats(b[0], b[1], b[2]));
        fn.setInt(org.apache.pdfbox.cos.COSName.N, 1);
        org.apache.pdfbox.pdmodel.graphics.shading.PDShadingType2 sh =
                new org.apache.pdfbox.pdmodel.graphics.shading.PDShadingType2(new org.apache.pdfbox.cos.COSDictionary());
        sh.setShadingType(org.apache.pdfbox.pdmodel.graphics.shading.PDShading.SHADING_TYPE2);
        sh.setColorSpace(org.apache.pdfbox.pdmodel.graphics.color.PDDeviceRGB.INSTANCE);
        sh.setCoords(floats((float) x0, (float) y0, (float) x1, (float) y1));
        sh.setFunction(org.apache.pdfbox.pdmodel.common.function.PDFunction.create(fn));
        org.apache.pdfbox.cos.COSArray ext = new org.apache.pdfbox.cos.COSArray();
        ext.add(org.apache.pdfbox.cos.COSBoolean.TRUE);
        ext.add(org.apache.pdfbox.cos.COSBoolean.TRUE);
        sh.setExtend(ext);
        alpha(1, 1);
        cs.saveGraphicsState();
        cs.addRect((float) x, (float) y, (float) w, (float) h);
        cs.clip();
        cs.shadingFill(sh);
        cs.restoreGraphicsState();
    }

    /**
     * 今の切り抜きの範囲を、(x0, y0) から (x1, y1) へ、位置 pos (0〜1) の色 cols を順に直線でつなぐグラデーションで塗る
     * (PDF の軸方向のシェーディング。色が 3 つ以上なら関数をつなぐ)。座標は今の変換の中の座標。
     */
    public void fillAxialStops(double x0, double y0, double x1, double y1, double[] pos, Color[] cols) throws IOException {
        fillAxialStops(x0, y0, x1, y1, pos, cols, 1f);
    }

    /** 不透明度 a (0〜1、全体で一定) を付けて塗る。 */
    public void fillAxialStops(double x0, double y0, double x1, double y1, double[] pos, Color[] cols, float a) throws IOException {
        if (cols.length == 0) return;
        org.apache.pdfbox.pdmodel.graphics.shading.PDShadingType2 sh =
                new org.apache.pdfbox.pdmodel.graphics.shading.PDShadingType2(new org.apache.pdfbox.cos.COSDictionary());
        sh.setShadingType(org.apache.pdfbox.pdmodel.graphics.shading.PDShading.SHADING_TYPE2);
        sh.setColorSpace(org.apache.pdfbox.pdmodel.graphics.color.PDDeviceRGB.INSTANCE);
        sh.setCoords(floats((float) x0, (float) y0, (float) x1, (float) y1));
        sh.setFunction(org.apache.pdfbox.pdmodel.common.function.PDFunction.create(stopsFunction(pos, cols)));
        org.apache.pdfbox.cos.COSArray ext = new org.apache.pdfbox.cos.COSArray();
        ext.add(org.apache.pdfbox.cos.COSBoolean.TRUE);
        ext.add(org.apache.pdfbox.cos.COSBoolean.TRUE);
        sh.setExtend(ext);
        alpha(a, a);
        cs.shadingFill(sh);
    }

    /**
     * 放射状のグラデーション (Type 3) でクリップの範囲を塗る。焦点 (fx, fy) の半径 0 の円から、中心 (cx, cy) の半径 r の円へ。
     * 円の外は最後の色、内は最初の色のまま延ばす。
     */
    public void fillRadialStops(double fx, double fy, double cx, double cy, double r, double[] pos, Color[] cols) throws IOException {
        fillRadialStops(fx, fy, cx, cy, r, pos, cols, 1f);
    }

    /** 不透明度 a (0〜1、全体で一定) を付けて塗る。 */
    public void fillRadialStops(double fx, double fy, double cx, double cy, double r, double[] pos, Color[] cols, float a) throws IOException {
        if (cols.length == 0 || !(r > 0)) return;
        org.apache.pdfbox.pdmodel.graphics.shading.PDShadingType3 sh =
                new org.apache.pdfbox.pdmodel.graphics.shading.PDShadingType3(new org.apache.pdfbox.cos.COSDictionary());
        sh.setShadingType(org.apache.pdfbox.pdmodel.graphics.shading.PDShading.SHADING_TYPE3);
        sh.setColorSpace(org.apache.pdfbox.pdmodel.graphics.color.PDDeviceRGB.INSTANCE);
        sh.setCoords(floats((float) fx, (float) fy, 0, (float) cx, (float) cy, (float) r));
        sh.setFunction(org.apache.pdfbox.pdmodel.common.function.PDFunction.create(stopsFunction(pos, cols)));
        org.apache.pdfbox.cos.COSArray ext = new org.apache.pdfbox.cos.COSArray();
        ext.add(org.apache.pdfbox.cos.COSBoolean.TRUE);
        ext.add(org.apache.pdfbox.cos.COSBoolean.TRUE);
        sh.setExtend(ext);
        alpha(a, a);
        cs.shadingFill(sh);
    }

    /** 色の位置 pos (0〜1) と色 cols の区分の関数 (Type 2 を Type 3 でつなぐ)。 */
    private org.apache.pdfbox.cos.COSDictionary stopsFunction(double[] pos, Color[] cols) {
        org.apache.pdfbox.cos.COSDictionary fn;
        if (cols.length == 1) {
            fn = type2(cols[0], cols[0]);
        } else {
            fn = new org.apache.pdfbox.cos.COSDictionary();
            fn.setInt(org.apache.pdfbox.cos.COSName.FUNCTION_TYPE, 3);
            fn.setItem(org.apache.pdfbox.cos.COSName.DOMAIN, floats(0, 1));
            org.apache.pdfbox.cos.COSArray fns = new org.apache.pdfbox.cos.COSArray(), bounds = new org.apache.pdfbox.cos.COSArray(), enc = new org.apache.pdfbox.cos.COSArray();
            // 最初の位置より前と最後の位置より後は、端の色のまま
            java.util.List<Double> p = new java.util.ArrayList<>();
            java.util.List<Color> c = new java.util.ArrayList<>();
            if (pos[0] > 0) { p.add(0.0); c.add(cols[0]); }
            for (int i = 0; i < cols.length; i++) { p.add(Math.max(0, Math.min(1, pos[i]))); c.add(cols[i]); }
            if (pos[pos.length - 1] < 1) { p.add(1.0); c.add(cols[cols.length - 1]); }
            for (int i = 0; i + 1 < c.size(); i++) {
                fns.add(type2(c.get(i), c.get(i + 1)));
                if (i > 0) bounds.add(new org.apache.pdfbox.cos.COSFloat(p.get(i).floatValue()));
                enc.add(new org.apache.pdfbox.cos.COSFloat(0));
                enc.add(new org.apache.pdfbox.cos.COSFloat(1));
            }
            if (fns.size() == 1) fn = (org.apache.pdfbox.cos.COSDictionary) fns.getObject(0);
            else {
                fn.setItem(org.apache.pdfbox.cos.COSName.FUNCTIONS, fns);
                fn.setItem(org.apache.pdfbox.cos.COSName.BOUNDS, bounds);
                fn.setItem(org.apache.pdfbox.cos.COSName.ENCODE, enc);
            }
        }
        return fn;
    }

    private org.apache.pdfbox.cos.COSDictionary type2(Color c0, Color c1) {
        org.apache.pdfbox.cos.COSDictionary fn = new org.apache.pdfbox.cos.COSDictionary();
        fn.setInt(org.apache.pdfbox.cos.COSName.FUNCTION_TYPE, 2);
        fn.setItem(org.apache.pdfbox.cos.COSName.DOMAIN, floats(0, 1));
        float[] a = rgb(c0).getRGBColorComponents(null), b = rgb(c1).getRGBColorComponents(null);
        fn.setItem(org.apache.pdfbox.cos.COSName.C0, floats(a[0], a[1], a[2]));
        fn.setItem(org.apache.pdfbox.cos.COSName.C1, floats(b[0], b[1], b[2]));
        fn.setInt(org.apache.pdfbox.cos.COSName.N, 1);
        return fn;
    }

    private static org.apache.pdfbox.cos.COSArray floats(float... v) {
        org.apache.pdfbox.cos.COSArray arr = new org.apache.pdfbox.cos.COSArray();
        for (float f : v) arr.add(new org.apache.pdfbox.cos.COSFloat(f));
        return arr;
    }

    /** ページのタイルの名前 (色と向きごと)。 */
    private final Map<String, org.apache.pdfbox.cos.COSName> tiles = new java.util.HashMap<>();

    /**
     * Excel が倍率 100% の点線に使うタイル (0.96pt 四方、8 画素。横線は左半分、縦線は上半分が線の色で残りは白。用紙に固定) で矩形を塗る。
     */
    public void fillDotTile(double x, double y, double w, double h, Color c, boolean horizontal) throws IOException {
        Color col = c == null ? Color.BLACK : rgb(c);
        long bits = 0;
        for (int r = 0; r < 8; r++) for (int k = 0; k < 8; k++) if (horizontal ? k < 4 : r < 4) bits |= 1L << (r * 8 + k);
        fillTile(x, y, w, h, bits, col, Color.WHITE);
    }

    /**
     * 網掛けの模様 (8 × 8 画素、ビット r × 8 + c が 1 なら前景色) で矩形を塗る。模様の 1 画素は 0.12pt (600 dpi) で、
     * 周期は用紙の左上に揃える (Excel の PDF と同じ)。
     */
    public void fillPatternTile(double x, double y, double w, double h, long bits, Color fg, Color bg) throws IOException {
        fillTile(x, y, w, h, bits, rgb(fg), rgb(bg));
    }

    private void fillTile(double x, double y, double w, double h, long bits, Color fg, Color bg) throws IOException {
        cs.saveGraphicsState();
        alpha(1, 1);   // 直前の半透明の塗りの透明度を受け継がない
        cs.setNonStrokingColor(pattern(bits, fg, bg, 0.12, 0, 0));
        cs.addRect((float) x, (float) y, (float) w, (float) h);
        cs.fill();
        cs.restoreGraphicsState();
    }

    /**
     * 図形を網掛けの模様で塗る。
     *
     * @param cell 模様の 1 マスの大きさ (用紙上の pt)
     * @param ox   模様の格子の基準の x (用紙の左端から、pt)
     * @param oy   模様の格子の基準の y (用紙の上端から、pt)
     */
    public void fillPattern(Shape s, long bits, Color fg, Color bg, double cell, double ox, double oy) throws IOException {
        if (fg == null || bg == null) return;
        cs.saveGraphicsState();
        alpha(1, 1);
        cs.setNonStrokingColor(pattern(bits, rgb(fg), rgb(bg), cell, ox, oy));
        path(s);
        if (windingEvenOdd(s)) cs.fillEvenOdd(); else cs.fill();
        cs.restoreGraphicsState();
    }

    /** 8 × 8 の模様のタイル (用紙の座標で、格子の左上が (ox, oy) に揃う)。 */
    private org.apache.pdfbox.pdmodel.graphics.color.PDColor pattern(long bits, Color fg, Color bg, double cell, double ox, double oy) throws IOException {
        String key = Long.toHexString(bits) + "/" + fg.getRGB() + "/" + bg.getRGB() + "/" + cell + "/" + ox + "/" + oy;
        org.apache.pdfbox.cos.COSName name = tiles.get(key);
        if (name == null) {
            java.awt.image.BufferedImage im = new java.awt.image.BufferedImage(8, 8, java.awt.image.BufferedImage.TYPE_INT_RGB);
            for (int r = 0; r < 8; r++) for (int k = 0; k < 8; k++)
                im.setRGB(k, r, (bits >>> (r * 8 + k) & 1) != 0 ? fg.getRGB() : bg.getRGB());
            PDImageXObject xo = org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory.createFromImage(doc, im);
            org.apache.pdfbox.pdmodel.PDResources res = new org.apache.pdfbox.pdmodel.PDResources();
            org.apache.pdfbox.cos.COSName xn = res.add(xo);
            org.apache.pdfbox.pdmodel.graphics.pattern.PDTilingPattern t = new org.apache.pdfbox.pdmodel.graphics.pattern.PDTilingPattern();
            t.setPaintType(1);
            t.setTilingType(2);
            t.setBBox(new org.apache.pdfbox.pdmodel.common.PDRectangle(0, 0, 8, 8));
            t.setXStep(8);
            t.setYStep(8);
            // 模様の空間は用紙の既定の座標 (左下が原点、上向き)。1 行目が格子の上端に来るよう、上端を pageH - oy に揃える
            t.setMatrix(new java.awt.geom.AffineTransform(cell, 0, 0, cell, ox, pageH - oy));
            t.setResources(res);
            try (java.io.OutputStream os = t.getContentStream().createOutputStream()) {
                os.write(("q 8 0 0 8 0 0 cm /" + xn.getName() + " Do Q").getBytes(java.nio.charset.StandardCharsets.US_ASCII));
            }
            name = page.getResources().add(t);
            tiles.put(key, name);
        }
        return new org.apache.pdfbox.pdmodel.graphics.color.PDColor(name, new org.apache.pdfbox.pdmodel.graphics.color.PDPattern(null));
    }

    /** 透明度を除いた色 (透明度はグラフィックス状態で指定する)。 */
    private static Color rgb(Color c) {
        return c.getAlpha() == 255 ? c : new Color(c.getRed(), c.getGreen(), c.getBlue());
    }

    /** 線分を描く。 */
    public void line(double x1, double y1, double x2, double y2, Stroke st, Color c) throws IOException {
        if (c == null) return;
        applyStroke(st, c);
        cs.moveTo((float) x1, (float) y1);
        cs.lineTo((float) x2, (float) y2);
        cs.stroke();
    }

    /** 線のスタイルと色を設定する。 */
    private void applyStroke(Stroke st, Color c) throws IOException {
        alpha(1, alphaOf(c));
        cs.setStrokingColor(rgb(c));
        cs.setLineWidth(st.width);
        cs.setLineCapStyle(st.cap);
        cs.setLineJoinStyle(st.join);
        if (st.dash != null && st.dash.length > 0) cs.setLineDashPattern(st.dash, 0);
        else cs.setLineDashPattern(new float[0], 0);
    }

    /** 図形を塗りつぶす (図形の塗りつぶし規則に従う)。 */
    public void fill(Shape s, Color c) throws IOException {
        if (c == null || c.getAlpha() == 0 || !finite(s)) return;
        alpha(alphaOf(c), 1);
        cs.setNonStrokingColor(rgb(c));
        path(s);
        if (windingEvenOdd(s)) cs.fillEvenOdd(); else cs.fill();
    }

    private static boolean windingEvenOdd(Shape s) {
        return s.getPathIterator(null).getWindingRule() == PathIterator.WIND_EVEN_ODD;
    }

    /** 図形の輪郭を描く。 */
    public void stroke(Shape s, Stroke st, Color c) throws IOException {
        if (c == null || c.getAlpha() == 0 || st == null || st.width <= 0 || !finite(s)) return;
        applyStroke(st, c);
        path(s);
        cs.stroke();
    }

    /**
     * 図形の座標がすべて有限か。大きさ 0 の図形の調整値の計算などで NaN になった図形は描かない
     * (書きかけのパスを残さないよう、書き出す前に調べる)。
     */
    private static boolean finite(Shape s) {
        float[] p = new float[6];
        for (PathIterator it = s.getPathIterator(null); !it.isDone(); it.next()) {
            int n = switch (it.currentSegment(p)) {
                case PathIterator.SEG_MOVETO, PathIterator.SEG_LINETO -> 2;
                case PathIterator.SEG_QUADTO -> 4;
                case PathIterator.SEG_CUBICTO -> 6;
                default -> 0;
            };
            for (int i = 0; i < n; i++) if (!Float.isFinite(p[i])) return false;
        }
        return true;
    }

    /** 図形をパスとして書き出す (2 次ベジエは 3 次に変換する)。 */
    private void path(Shape s) throws IOException {
        PathIterator it = s.getPathIterator(null);
        float[] p = new float[6];
        float cx = 0, cy = 0;
        while (!it.isDone()) {
            switch (it.currentSegment(p)) {
                case PathIterator.SEG_MOVETO -> { cs.moveTo(p[0], p[1]); cx = p[0]; cy = p[1]; }
                case PathIterator.SEG_LINETO -> { cs.lineTo(p[0], p[1]); cx = p[0]; cy = p[1]; }
                case PathIterator.SEG_QUADTO -> {
                    float c1x = cx + 2f / 3f * (p[0] - cx), c1y = cy + 2f / 3f * (p[1] - cy);
                    float c2x = p[2] + 2f / 3f * (p[0] - p[2]), c2y = p[3] + 2f / 3f * (p[1] - p[3]);
                    cs.curveTo(c1x, c1y, c2x, c2y, p[2], p[3]);
                    cx = p[2]; cy = p[3];
                }
                case PathIterator.SEG_CUBICTO -> { cs.curveTo(p[0], p[1], p[2], p[3], p[4], p[5]); cx = p[4]; cy = p[5]; }
                case PathIterator.SEG_CLOSE -> cs.closePath();
                default -> { }
            }
            it.next();
        }
    }

    /** 画像を矩形に描く (左上原点)。 */
    public void image(PDImageXObject img, double x, double y, double w, double h, float alpha) throws IOException {
        save();
        alpha(alpha, 1);
        cs.drawImage(img, new Matrix((float) w, 0, 0, (float) -h, (float) x, (float) (y + h)));
        restore();
    }

    /**
     * 文字列をベースラインの位置に描く。グリフを持つフォントごとに分けて書き出す。
     *
     * @param hScale        横方向の伸縮 (1 = 標準)
     * @param letterSpacing 追加の文字間隔 (pt)
     */
    public void text(String s, double x, double baseline, FontManager.Spec spec, float size, Color color,
                     float hScale, float letterSpacing) throws IOException {
        if (s.isEmpty() || size <= 0) return;
        s = ComplexText.cluster(s);
        if (mirrorText) {
            // 右から左のシート: ページ全体の左右反転を、文字列の中心を軸に打ち消す (字形は読める向きのまま)
            double w = 0;
            for (int k = 0; k < s.length(); ) {
                int cp = s.codePointAt(k);
                k += Character.charCount(cp);
                w += spec.charWidth(cp == '\n' || cp == '\r' || cp == '\t' ? ' ' : cp, size, fonts) * hScale + letterSpacing;
            }
            save();
            transform(flipAbout(x + w / 2));
            mirrorText = false;
            try {
                text(s, x, baseline, spec, size, color, hScale, letterSpacing);
            } finally {
                mirrorText = true;
                restore();
            }
            return;
        }
        alpha(alphaOf(color), alphaOf(color));
        cs.setNonStrokingColor(rgb(color));
        if (spec.bold) {
            cs.setStrokingColor(rgb(color));
            cs.setLineWidth(size * 0.035f);
            cs.setLineJoinStyle(1);
            cs.setLineDashPattern(new float[0], 0);
        }
        List<Object[]> shapedAt = null;
        cs.beginText();
        double pos = x;
        int i = 0;
        FontManager.Face current = null;
        StringBuilder run = new StringBuilder();
        double runStart = x;
        while (i < s.length()) {
            int cp = s.codePointAt(i);
            i += Character.charCount(cp);
            if (cp == '\n' || cp == '\r' || cp == '\t') cp = ' ';
            FontManager.Face f = spec.faceFor(cp, fonts);
            if (ComplexText.isCluster(cp)) {
                // 複雑な文字体系のまとまり: 文字の出力の後で、整形した輪郭を塗る
                if (run.length() > 0) { emit(current, run, runStart, baseline, size, hScale, spec, letterSpacing); run.setLength(0); }
                if (shapedAt == null) shapedAt = new ArrayList<>();
                shapedAt.add(new Object[] {f, cp, pos});
                pos += spec.charWidth(cp, size, fonts) * hScale + letterSpacing;
                runStart = pos;
                current = null;
                continue;
            }
            if (!f.hasGlyph(cp) || FontManager.isPad(cp)) {
                // グリフを持つフォントがない: 送り幅だけ空ける
                if (run.length() > 0) { emit(current, run, runStart, baseline, size, hScale, spec, letterSpacing); run.setLength(0); }
                pos += spec.charWidth(cp, size, fonts) * hScale + letterSpacing;
                runStart = pos;
                current = null;
                continue;
            }
            if (f != current && run.length() > 0) {
                emit(current, run, runStart, baseline, size, hScale, spec, letterSpacing);
                run.setLength(0);
            }
            if (run.length() == 0) runStart = pos;
            current = f;
            run.appendCodePoint(cp);
            pos += spec.charWidth(cp, size, fonts) * hScale + letterSpacing;
        }
        if (run.length() > 0) emit(current, run, runStart, baseline, size, hScale, spec, letterSpacing);
        cs.endText();
        if (shapedAt != null) {
            for (Object[] o : shapedAt) {
                FontManager.Face f = (FontManager.Face) o[0];
                AffineTransform at = new AffineTransform(size * hScale, 0, spec.synthItalic(f) ? -0.21 * size : 0, size, (Double) o[2], baseline);
                Shape g = at.createTransformedShape(f.shaped((Integer) o[1]).path);
                fill(g, color);
                if (spec.synthBold(f)) stroke(g, new Stroke(size * 0.035f), color);
            }
        }
    }

    /** 1 つのフォントの文字列を書き出す (合成太字・合成斜体、グリフのピクセル格子への吸着を含む)。 */
    private void emit(FontManager.Face f, StringBuilder run, double x, double baseline, float size, float hScale,
                      FontManager.Spec spec, float letterSpacing) throws IOException {
        cs.setFont(fonts.pdfFont(f), size);
        boolean sb = spec.synthBold(f);
        cs.setRenderingMode(sb ? RenderingMode.FILL_STROKE : RenderingMode.FILL);
        float shear = spec.synthItalic(f) ? 0.21f : 0f;
        cs.setTextMatrix(new Matrix(hScale, 0, shear, -1, (float) x, (float) baseline));
        float extra = sb ? 0.02f * size : 0f;
        float cspace = (extra * hScale + letterSpacing) / hScale;
        cs.setCharacterSpacing(cspace);
        try {
            if (FontManager.gridActive() && hScale == 1f) showSnapped(f, run, size, hScale, spec, extra);
            else cs.showText(run.toString());
        } catch (IllegalArgumentException e) {
            // グリフの符号化の問題: 1 文字ずつ描き、失敗した文字は飛ばす
            for (int i = 0; i < run.length(); ) {
                int cp = run.codePointAt(i);
                i += Character.charCount(cp);
                try {
                    cs.showText(new String(Character.toChars(cp)));
                } catch (IllegalArgumentException ignored) {
                    Warnings.add(String.format("文字「%s」(U+%04X) をフォント「%s」で描けませんでした", new String(Character.toChars(cp)), cp, f.key), null);
                }
            }
        }
        cs.setCharacterSpacing(0);
    }

    /**
     * 各グリフの送り幅をデバイスのピクセル格子に丸めて ({@link FontManager#snap}) 文字列を書き出す。
     * フォント本来の送り幅との差を TJ の位置調整で与える。
     */
    private void showSnapped(FontManager.Face f, StringBuilder run, float size, float hScale, FontManager.Spec spec,
                             float boldExtra) throws IOException {
        List<Object> parts = new ArrayList<>();
        StringBuilder chunk = new StringBuilder();
        for (int i = 0; i < run.length(); ) {
            int cp = run.codePointAt(i);
            i += Character.charCount(cp);
            chunk.appendCodePoint(cp);
            float natural = f.advance(cp) * size + boldExtra;
            float snapped = spec.charWidth(cp, size, fonts);
            if (Math.abs(snapped - natural) > 1e-4) {
                parts.add(chunk.toString());
                chunk.setLength(0);
                // TJ の数値はテキスト空間の 1/1000 単位で、正の値は左へ動かす
                parts.add((natural - snapped) / size * 1000f);
            }
        }
        if (chunk.length() > 0) parts.add(chunk.toString());
        cs.showTextWithPositioning(parts.toArray());
    }

    /** 1 文字を、ベースラインの原点を中心に angle (ラジアン) 回転して描く。 */
    public void glyph(String ch, double x, double baseline, double angle, FontManager.Spec spec, float size, Color color)
            throws IOException {
        if (mirrorText) {
            save();
            transform(flipAbout(x));
            mirrorText = false;
            try {
                glyph(ch, x, baseline, -angle, spec, size, color);
            } finally {
                mirrorText = true;
                restore();
            }
            return;
        }
        save();
        AffineTransform at = AffineTransform.getTranslateInstance(x, baseline);
        at.rotate(angle);
        transform(at);
        text(ch, 0, 0, spec, size, color, 1, 0);
        restore();
    }

    /** 矩形を作る。 */
    public static Rectangle2D rect(double x, double y, double w, double h) {
        return new Rectangle2D.Double(x, y, w, h);
    }

    /** コンテンツストリームを閉じる。 */
    @Override
    public void close() throws IOException {
        cs.close();
    }
}
