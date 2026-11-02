package xlsx2pdf.render;

import java.awt.Color;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;

import xlsx2pdf.render.CondFormat.Edge;

/** 縦横の罫線を、Excel と同じくプリンタの画素 (600 dpi、1 画素 0.12pt) の帯で描く。 */
final class BorderLines {
    private BorderLines() {}

    private static long px(double base, double scale) { return (long) Math.floor(base * scale + 0.5 + 1e-9); }

    /** 破線の区間 (プリンタの画素、描く・空けるの順。倍率 100%)。 */
    private static int[] dashPx(BorderStyle s) {
        return switch (s) {
            case DASHED -> new int[] {24, 8};
            case DASH_DOT, MEDIUM_DASH_DOT -> new int[] {72, 24, 24, 24};
            case DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT -> new int[] {72, 24, 24, 24, 24, 24};
            case MEDIUM_DASHED -> new int[] {72, 24};
            default -> null;
        };
    }

    /**
     * 縦横の罫線を Excel のプリンタの画素で描く。u は 1 画素。帯の太さは round(8 / 16 / 24 × 倍率)、長さは区間 + 1 画素。
     * 極細はどの倍率でも 1 画素の帯と 0.14pt の線。二重線は太線の帯の両端に細線の太さの帯。破線の区間は round(基準 × 倍率) で、周期は用紙の座標に固定。
     */
    static void draw(PdfCanvas cv, Edge e, double x0, double y0, double x1, double y1, double scale, double originX, double originY) throws IOException {
        double u = 0.12 / scale;
        BorderStyle bs = e.style();
        Color col = e.color() != null ? e.color() : Color.BLACK;
        boolean hz = y0 == y1;
        long thin = px(8, scale);
        double adj = thin % 2 == 1 ? 0.5 * u : 0;
        double c = (hz ? y0 : x0) - adj;
        double a0 = Math.min(hz ? x0 : y0, hz ? x1 : y1) - adj, a1 = Math.max(hz ? x0 : y0, hz ? x1 : y1) - adj + u;
        if (bs == BorderStyle.HAIR) {
            band(cv, hz, c, u, a0, a1, col);
            hairline(cv, scale, hz, c + 0.5 * u, a0 + 0.5 * u, a1 - 0.5 * u, col);
            return;
        }
        if (bs == BorderStyle.DOUBLE) {
            long t = px(24, scale);
            double st = c - (t / 2) * u;
            band(cv, hz, st, thin * u, a0, a1, col);
            band(cv, hz, st + (t - thin) * u, thin * u, a0, a1, col);
            return;
        }
        int base = switch (bs) {
            case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT -> 16;
            case THICK -> 24;
            default -> 8;
        };
        long w = px(base, scale);
        double st = c - (w / 2) * u;
        boolean line = scale == 1 && base == 8;
        int[] pat = dashPx(bs);
        long[] seg = null;
        if (pat != null) {
            seg = new long[pat.length];
            for (int i = 0; i < pat.length; i++) seg[i] = px(pat[i], scale);
        } else if (bs == BorderStyle.DOTTED) {
            if (scale == 1) {
                if (hz) cv.fillDotTile(a0, st, a1 - a0, w * u, col, true);
                else cv.fillDotTile(st, a0, w * u, a1 - a0, col, false);
                return;
            }
            seg = new long[] {w, w};
            line = false;
        }
        if (seg == null) {
            band(cv, hz, st, w * u, a0, a1, col);
            if (line) hairline(cv, scale, hz, st + 0.5 * u, a0 + 0.5 * u, a1 - 0.5 * u, col);
            return;
        }
        long period = 0;
        for (long v : seg) period += v;
        double org = hz ? originX : originY;   // シートの座標 0 の用紙上の位置
        long pa0 = Math.round((org + a0 * scale) / 0.12), pa1 = Math.round((org + a1 * scale) / 0.12);
        double back = a0 - pa0 * u;                                        // 画素 → シートの座標
        long k = Math.floorDiv(pa0, period) * period;
        while (k < pa1) {
            long pos = k;
            for (int i = 0; i < seg.length; i += 2) {
                long b0 = Math.max(pos, pa0), b1 = Math.min(pos + seg[i], pa1);
                if (b1 > b0) {
                    band(cv, hz, st, w * u, back + b0 * u, back + b1 * u, col);
                    if (line) hairline(cv, scale, hz, st + 0.5 * u, back + b0 * u + 0.5 * u, back + b1 * u - 0.5 * u, col);
                }
                pos += seg[i] + (i + 1 < seg.length ? seg[i + 1] : 0);
            }
            k += period;
        }
    }

    /** 帯 (太さ方向 p から d、長さ方向 a0 から a1)。 */
    private static void band(PdfCanvas cv, boolean hz, double p, double d, double a0, double a1, Color col) throws IOException {
        if (hz) cv.fillRect(a0, p, a1 - a0, d, col);
        else cv.fillRect(p, a0, d, a1 - a0, col);
    }

    /** Excel が帯に重ねる 0.14pt の線。 */
    private static void hairline(PdfCanvas cv, double scale, boolean hz, double p, double a0, double a1, Color col) throws IOException {
        PdfCanvas.Stroke h = new PdfCanvas.Stroke((float) (0.14 / scale));
        h.cap = 2;
        if (hz) cv.line(a0, p, a1, p, h, col);
        else cv.line(p, a0, p, a1, h, col);
    }
}
