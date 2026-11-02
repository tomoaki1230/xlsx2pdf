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
    /** ブック全体で共有する画像のキャッシュ。 */
    private final Map<Object, PDImageXObject> imageCache;

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
     *
     * @param imageCache ブック全体で共有する画像のキャッシュ
     */
    public PdfCanvas(PDDocument doc, PDPage page, FontManager fonts, Map<Object, PDImageXObject> imageCache) throws IOException {
        this.doc = doc;
        this.fonts = fonts;
        this.imageCache = imageCache;
        this.page = page;
        if (page.getResources() == null) page.setResources(new org.apache.pdfbox.pdmodel.PDResources());
        this.cs = new PDPageContentStream(doc, page);
        float h = page.getMediaBox().getHeight();
        this.pageH = h;
        cs.transform(new Matrix(1, 0, 0, -1, 0, h));
    }

    /** キャッシュした画像を返す (なければ null)。 */
    public PDImageXObject cachedImage(Object key) { return imageCache.get(key); }

    /** 画像をキャッシュする。 */
    public void cacheImage(Object key, PDImageXObject img) { imageCache.put(key, img); }

    /** グラフィックス状態を保存する。 */
    public void save() throws IOException { cs.saveGraphicsState(); }

    /** グラフィックス状態を戻す。 */
    public void restore() throws IOException { cs.restoreGraphicsState(); }

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

    /** ページのタイルの名前 (色と向きごと)。 */
    private final Map<String, org.apache.pdfbox.cos.COSName> tiles = new java.util.HashMap<>();

    /**
     * Excel が倍率 100% の点線に使うタイル (0.96pt 四方、8 画素。横線は左半分、縦線は上半分が線の色で残りは白。用紙に固定) で矩形を塗る。
     */
    public void fillDotTile(double x, double y, double w, double h, Color c, boolean horizontal) throws IOException {
        Color col = c == null ? Color.BLACK : rgb(c);
        String key = col.getRGB() + (horizontal ? "h" : "v");
        org.apache.pdfbox.cos.COSName name = tiles.get(key);
        if (name == null) {
            java.awt.image.BufferedImage im = new java.awt.image.BufferedImage(8, 8, java.awt.image.BufferedImage.TYPE_INT_RGB);
            for (int i = 0; i < 8; i++) for (int j = 0; j < 8; j++)
                im.setRGB(i, j, (horizontal ? i < 4 : j < 4) ? col.getRGB() : 0xFFFFFF);
            PDImageXObject xo = org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory.createFromImage(doc, im);
            org.apache.pdfbox.pdmodel.PDResources res = new org.apache.pdfbox.pdmodel.PDResources();
            org.apache.pdfbox.cos.COSName xn = res.add(xo);
            org.apache.pdfbox.pdmodel.graphics.pattern.PDTilingPattern t = new org.apache.pdfbox.pdmodel.graphics.pattern.PDTilingPattern();
            t.setPaintType(1);
            t.setTilingType(2);
            t.setBBox(new org.apache.pdfbox.pdmodel.common.PDRectangle(0, 0, 0.96f, 0.96f));
            t.setXStep(0.96f);
            t.setYStep(0.96f);
            t.setMatrix(new java.awt.geom.AffineTransform(1, 0, 0, 1, 0, pageH));
            t.setResources(res);
            try (java.io.OutputStream os = t.getContentStream().createOutputStream()) {
                os.write(("q 0.96 0 0 0.96 0 0 cm /" + xn.getName() + " Do Q").getBytes(java.nio.charset.StandardCharsets.US_ASCII));
            }
            name = page.getResources().add(t);
            tiles.put(key, name);
        }
        cs.saveGraphicsState();
        cs.setNonStrokingColor(new org.apache.pdfbox.pdmodel.graphics.color.PDColor(name, new org.apache.pdfbox.pdmodel.graphics.color.PDPattern(null)));
        cs.addRect((float) x, (float) y, (float) w, (float) h);
        cs.fill();
        cs.restoreGraphicsState();
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
        if (c == null || c.getAlpha() == 0) return;
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
        if (c == null || c.getAlpha() == 0 || st == null || st.width <= 0) return;
        applyStroke(st, c);
        path(s);
        cs.stroke();
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
        alpha(alphaOf(color), alphaOf(color));
        cs.setNonStrokingColor(rgb(color));
        if (spec.bold) {
            cs.setStrokingColor(rgb(color));
            cs.setLineWidth(size * 0.035f);
            cs.setLineJoinStyle(1);
            cs.setLineDashPattern(new float[0], 0);
        }
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
            if (!f.hasGlyph(cp)) {
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
                    // 飛ばす
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
