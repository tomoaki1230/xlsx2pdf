package xlsx2pdf.render;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Locale;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.hemf.usermodel.HemfPicture;
import org.apache.poi.hwmf.usermodel.HwmfPicture;

/** 埋め込まれた画像データを PDF の画像に変換する。 */
public final class Images {
    private Images() {}

    /** メタファイル (EMF/WMF) をラスタライズするときの解像度 (dpi)。 */
    private static final double METAFILE_DPI = 200;
    /** 1 枚の画像の画素数の上限 (0 以下は無制限)。{@link xlsx2pdf.Converter.Options#maxImagePixels}。 */
    private static final ThreadLocal<Long> MAX_PIXELS = ThreadLocal.withInitial(() -> 50_000_000L);

    /** 画像の画素数の上限を設定する (0 以下は無制限)。 */
    public static void setMaxPixels(long n) {
        MAX_PIXELS.set(n);
    }

    /** 幅 w × 高さ h の画像が上限内か (上限を超える画像は描かない)。 */
    static boolean pixelsAllowed(long w, long h) {
        long m = MAX_PIXELS.get();
        return w > 0 && h > 0 && (m <= 0 || w * h <= m);
    }

    /**
     * 画像をデコードせずに大きさを読む (ヘッダーだけ)。読めなければ null。
     * 小さなファイルが巨大な画像を名乗って、デコードでメモリを使い切るのを防ぐため、デコードの前に確かめる。
     */
    static int[] headerSize(byte[] data) {
        try (javax.imageio.stream.ImageInputStream iis = ImageIO.createImageInputStream(new ByteArrayInputStream(data))) {
            if (iis == null) return null;
            java.util.Iterator<javax.imageio.ImageReader> it = ImageIO.getImageReaders(iis);
            if (!it.hasNext()) return null;
            javax.imageio.ImageReader r = it.next();
            try {
                r.setInput(iis, true, true);
                return new int[] {r.getWidth(0), r.getHeight(0)};
            } finally {
                r.dispose();
            }
        } catch (IOException | RuntimeException e) {
            return null;
        }
    }

    /** 上限内なら画像をデコードする (上限を超えるか読めなければ null)。 */
    static BufferedImage decodeLimited(byte[] data) throws IOException {
        int[] wh = headerSize(data);
        if (wh == null || !pixelsAllowed(wh[0], wh[1])) return null;
        return ImageIO.read(new ByteArrayInputStream(data));
    }

    /** ラスタライズする画像の 1 辺の上限 (ピクセル)。 */
    private static final int METAFILE_MAX_PX = 4000;

    /**
     * 画像データがメタファイル (EMF/WMF、圧縮版の EMZ/WMZ を含む) かを判定する。
     *
     * @param data 画像データ
     * @param name パーツ名 (拡張子の判定に使う。null 可)
     */
    public static boolean isMetafile(byte[] data, String name) {
        String n = name == null ? "" : name.toLowerCase(Locale.ROOT);
        if (n.endsWith(".emf") || n.endsWith(".wmf") || n.endsWith(".emz") || n.endsWith(".wmz")) return true;
        if (isGzip(data)) {
            byte[] u = gunzip(data);
            return u != data && isMetafile(u, null);
        }
        if (data.length > 44 && data[0] == 1 && data[1] == 0 && data[2] == 0 && data[3] == 0 && isEmf(data)) return true;
        // 配置可能 WMF のマジックナンバー D7 CD C6 9A
        return data.length > 4 && (data[0] & 255) == 0xD7 && (data[1] & 255) == 0xCD && (data[2] & 255) == 0xC6
                && (data[3] & 255) == 0x9A;
    }

    /** gzip のマジックナンバー 1F 8B を持つか (EMZ・WMZ)。 */
    static boolean isGzip(byte[] data) {
        return data.length > 2 && (data[0] & 255) == 0x1F && (data[1] & 255) == 0x8B;
    }

    /** 展開後の大きさの上限 (バイト)。 */
    private static final int GUNZIP_MAX = 256 << 20;

    /** gzip で圧縮されていれば展開する (EMZ・WMZ)。圧縮されていない・展開できなければそのまま返す。 */
    static byte[] gunzip(byte[] data) {
        if (!isGzip(data)) return data;
        try (java.util.zip.GZIPInputStream in = new java.util.zip.GZIPInputStream(new ByteArrayInputStream(data));
             java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream()) {
            byte[] buf = new byte[65536];
            for (int n; (n = in.read(buf)) > 0; ) {
                out.write(buf, 0, n);
                if (out.size() > GUNZIP_MAX) return data;
            }
            return out.toByteArray();
        } catch (IOException e) {
            return data;
        }
    }

    /** EMF ヘッダーのシグネチャ (オフセット 40 の " EMF") を持つか。 */
    private static boolean isEmf(byte[] data) {
        return data.length > 44 && data[40] == ' ' && data[41] == 'E' && data[42] == 'M' && data[43] == 'F';
    }

    /**
     * PDF の画像を作る (同じキーは 1 回だけ作ってキャッシュする)。メタファイルは指定サイズでラスタライズする。
     *
     * @param key  キャッシュのキー
     * @param data 画像データ
     * @param name パーツ名 (形式の判定に使う)
     * @param wPt  描画する幅 (pt。メタファイルの解像度の決定に使う)
     * @param hPt  描画する高さ (pt)
     * @return PDF の画像。読めない形式なら null
     */
    public static PDImageXObject toPdf(RenderContext ctx, Object key, byte[] data, String name, double wPt, double hPt) {
        PDImageXObject cached = ctx.images.get(key);
        if (cached != null) return cached;
        PDImageXObject img = null;
        try {
            if (isMetafile(data, name)) {
                BufferedImage bi = renderMetafile(data, name, wPt, hPt, ctx.fonts);
                if (bi != null) img = LosslessFactory.createFromImage(ctx.doc, bi);
            } else if (data.length > 3 && (data[0] & 255) == 0xFF && (data[1] & 255) == 0xD8) {
                // JPEG はそのまま埋め込む。PDFBox が読めない JPEG (CMYK など) は RGB に変換して埋め込み直す
                int[] wh = headerSize(data);
                if (wh != null && !pixelsAllowed(wh[0], wh[1])) {
                    Warnings.add("画像「" + name + "」は画素数の上限を超えるため描きませんでした", null);
                    return null;
                }
                try {
                    img = JPEGFactory.createFromByteArray(ctx.doc, data);
                } catch (IOException | RuntimeException e) {
                    BufferedImage bi = decodeLimited(data);
                    if (bi != null) img = JPEGFactory.createFromImage(ctx.doc, toRgb(bi), 0.92f);
                }
            } else {
                BufferedImage bi = decodeLimited(data);
                if (bi != null) img = LosslessFactory.createFromImage(ctx.doc, bi);
                else Warnings.add("画像「" + name + "」は読めない形式か、画素数の上限を超えるため描きませんでした", null);
            }
        } catch (IOException | RuntimeException e) {
            Warnings.add("画像「" + name + "」を変換できませんでした", e);
            img = null;
        }
        if (img != null) ctx.images.put(key, img);
        return img;
    }

    /** 透過のない RGB 画像に変換する (透過部分は白)。 */
    private static BufferedImage toRgb(BufferedImage bi) {
        BufferedImage o = new BufferedImage(bi.getWidth(), bi.getHeight(), BufferedImage.TYPE_INT_RGB);
        Graphics2D g = o.createGraphics();
        g.drawImage(bi, 0, 0, java.awt.Color.WHITE, null);
        g.dispose();
        return o;
    }

    /**
     * ラスター画像のピクセルサイズを返す。
     *
     * @return {幅, 高さ}。読めなければ null
     */
    public static int[] size(byte[] data) {
        return headerSize(data);
    }

    /**
     * メタファイルを指定サイズ (pt) の画像にラスタライズする。
     *
     * @return 画像。描画に失敗したら null
     */
    static BufferedImage renderMetafile(byte[] data, String name, double wPt, double hPt, FontManager fonts) throws IOException {
        data = gunzip(data); // EMZ・WMZ
        int w = (int) Math.max(1, Math.min(METAFILE_MAX_PX, Math.round(wPt / 72 * METAFILE_DPI)));
        int h = (int) Math.max(1, Math.min(METAFILE_MAX_PX, Math.round(hPt / 72 * METAFILE_DPI)));
        BufferedImage bi = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = bi.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        try {
            if (isEmf(data) && fonts != null) {
                // 自前の EMF 描画 (POI の描画は崩れる)。何も描けなければ POI に任せる
                EmfRenderer er = new EmfRenderer(data, g, fonts);
                er.draw(w, h);
                if (er.drawn > 0) return bi;
                g.dispose();
                bi = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
                g = bi.createGraphics();
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            }
            if (isEmf(data)) {
                new HemfPicture(new ByteArrayInputStream(data)).draw(g, new Rectangle2D.Double(0, 0, w, h));
            } else {
                new HwmfPicture(new ByteArrayInputStream(data)).draw(g, new Rectangle2D.Double(0, 0, w, h));
            }
        } catch (RuntimeException | IOException e) {
            return null;
        } finally {
            g.dispose();
        }
        return bi;
    }

}
