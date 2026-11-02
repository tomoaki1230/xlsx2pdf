package xlsx2pdf.render;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

/** セルと図形で共通の、書式付き文字列 (ラン) と行分割。 */
public final class TextLayout {
    private TextLayout() {}

    /** 均等割り付けで、文字 cp と次の文字 next の間に間隔を配るか。 */
    static boolean distGapAfter(int cp, int next) {
        // 続く空白はまとめて 1 つの間隔に含める (空白の前には配らず、空白の後に配る。
        // 「名｜称」「称｜：」の 2 か所に同じ幅を配る)
        if (isSpace(next)) return false;
        return FontManager.isCjk(cp) || FontManager.isCjk(next) || isSpace(cp);
    }

    /** 行の文字を順に並べる (ランをまたいで)。 */
    static int[] codePoints(Line line) {
        StringBuilder sb = new StringBuilder();
        for (Run r : line.runs) sb.append(r.text);
        return sb.codePoints().toArray();
    }

    /** 均等割り付けで間隔を配る位置の数 (0 なら 1 つの塊として中央に置く)。 */
    public static int distGaps(Line line) {
        int[] cps = codePoints(line);
        int n = 0;
        for (int i = 0; i + 1 < cps.length; i++) if (distGapAfter(cps[i], cps[i + 1])) n++;
        return n;
    }

    /** 折り返しで空白として扱う文字か。 */
    private static boolean isSpace(int cp) {
        return cp == ' ' || cp == '\u3000';
    }

    /** 行頭に置かない文字 (禁則処理)。 */
    private static final String NO_START = "、。，．・：；？！ー）」』】〕〉》’”ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ々〻,.!?:;)]}％%";
    /** 行末に置かない文字 (禁則処理)。 */
    private static final String NO_END = "（「『【〔〈《‘“([{";

    /** 1 つの書式で書かれた文字列。 */
    public static final class Run {
        /** 文字列。 */
        public String text;
        /** フォント。 */
        public FontManager.Spec spec;
        /** フォントサイズ (pt)。 */
        public float size;
        /** 元のフォント名と大きさ (pt。SheetLayout の REWRAP_RICH・HeaderFooter の HF_CAL)。行分割で分かれたランにも引き継ぐ。 */
        public String fontName;
        public double ptSize;
        /** 文字色。 */
        public Color color;
        /** 下線 (0 なし、1 一重、2 二重)。 */
        public int underline;
        /** 取り消し線。 */
        public boolean strike;
        /** 上付き (1)・下付き (-1)。 */
        public int baselineShift;
        /** 追加の文字間隔 (pt)。 */
        public float spacing;

        /** ランを作る。 */
        public Run(String text, FontManager.Spec spec, float size, Color color) {
            this.text = text;
            this.spec = spec;
            this.size = size;
            this.color = color;
        }

        /** 書式を引き継いで、文字列だけ置き換えたランを返す。 */
        public Run copy(String t) {
            Run r = new Run(t, spec, size, color);
            r.underline = underline;
            r.strike = strike;
            r.baselineShift = baselineShift;
            r.spacing = spacing;
            r.fontName = fontName;
            r.ptSize = ptSize;
            return r;
        }

        /** 実際に描くサイズ (上付き・下付きは 66%)。 */
        float effSize() { return baselineShift != 0 ? size * 0.66f : size; }
    }

    /** 行分割した 1 行。 */
    public static final class Line {
        /** 行のラン。 */
        public final List<Run> runs = new ArrayList<>();
        /** 行の幅 (pt)。 */
        public float width;
        /** 空行の高さを決めるラン (空行でなければ null)。 */
        public Run blankBase;
        /** 行のアセント・ディセント (pt)。 */
        public float ascent, descent;
        /** 段落の最後の行 (改行または文字列の終わり) か。 */
        public boolean endsParagraph;
        /** 空白 (半角・全角) の数。 */
        public int spaces;
        /** 文字数。 */
        public int chars;
        /** 行末の空白 (半角・全角) の幅 (pt)。中央揃えの位置はこれを除いた幅で決める。 */
        public float trailSpace;
    }

    /** ランの幅 (pt。文字間隔を含む)。 */
    public static float width(Run r, FontManager fm) {
        float w = r.spec.width(r.text, r.effSize(), fm);
        if (r.spacing != 0) w += r.spacing * r.text.codePointCount(0, r.text.length());
        return w;
    }

    private static boolean noStart(int cp) {
        return NO_START.indexOf(cp) >= 0;
    }

    private static boolean noEnd(int cp) {
        return NO_END.indexOf(cp) >= 0;
    }

    /** prev と cp の間で改行してよいか (空白の後、または禁則に触れない CJK 文字の前後とハイフンの後)。 */
    private static boolean breakable(int prev, int cp) {
        if (isSpace(prev) || prev == '\t') return true;
        if (noStart(cp) || noEnd(prev)) return false;
        return FontManager.isCjk(cp) || FontManager.isCjk(prev) || prev == '-';
    }

    /** ランから展開した 1 文字と幅。 */
    private record Ch(int cp, Run run, float w) {}

    /**
     * ランを行に分割する。
     *
     * @param runs         ラン
     * @param maxWidth     折り返す幅 (pt)。0 以下なら折り返さない
     * @param wrap         折り返すか
     * @param keepNewlines 改行文字で改行するか (false なら改行文字を無視する)
     * @return 行の一覧 (少なくとも 1 行)
     */
    public static List<Line> layout(List<Run> runs, float maxWidth, boolean wrap, boolean keepNewlines, FontManager fm) {
        List<Ch> chars = new ArrayList<>();
        for (Run r : runs) {
            String t = r.text;
            for (int i = 0; i < t.length(); ) {
                int cp = t.codePointAt(i);
                i += Character.charCount(cp);
                if (cp == '\r') continue;
                if (cp == '\n' && !keepNewlines) continue;
                float w = cp == '\n' ? 0 : r.spec.charWidth(cp, r.effSize(), fm) + r.spacing;
                chars.add(new Ch(cp, r, w));
            }
        }
        List<List<Ch>> lines = new ArrayList<>();
        List<Boolean> paraEnd = new ArrayList<>();
        List<Ch> cur = new ArrayList<>();
        float curW = 0;
        int lastBreak = -1; // 改行できる位置 (cur の添字。行はその直前で終わる)
        for (Ch c : chars) {
            if (c.cp == '\n') {
                lines.add(cur);
                paraEnd.add(true);
                cur = new ArrayList<>();
                curW = 0;
                lastBreak = -1;
                continue;
            }
            if (wrap && maxWidth > 0 && !cur.isEmpty()) {
                int prev = cur.get(cur.size() - 1).cp;
                if (breakable(prev, c.cp)) lastBreak = cur.size();
                if (curW + c.w > maxWidth + 0.01 && !isSpace(c.cp)) {
                    List<Ch> next;
                    if (lastBreak > 0) {
                        next = new ArrayList<>(cur.subList(lastBreak, cur.size()));
                        cur = new ArrayList<>(cur.subList(0, lastBreak));
                    } else {
                        next = new ArrayList<>();
                    }
                    // 行末の空白は幅に数えない (後で除く)
                    lines.add(cur);
                    paraEnd.add(false);
                    // 次の行の先頭の空白は捨てる
                    while (!next.isEmpty() && isSpace(next.get(0).cp)) next.remove(0);
                    cur = next;
                    curW = 0;
                    for (Ch x : cur) curW += x.w;
                    lastBreak = -1;
                }
            }
            cur.add(c);
            curW += c.w;
        }
        lines.add(cur);
        paraEnd.add(true);

        List<Line> out = new ArrayList<>();
        for (int li = 0; li < lines.size(); li++) {
            List<Ch> l = lines.get(li);
            Line line = new Line();
            line.endsParagraph = paraEnd.get(li);
            int end = l.size();
            if (wrap) while (end > 0 && isSpace(l.get(end - 1).cp)) end--;
            // 同じランの文字をまとめ直す
            StringBuilder sb = new StringBuilder();
            Run rr = null;
            for (int i = 0; i < end; i++) {
                Ch c = l.get(i);
                if (c.run != rr && sb.length() > 0) {
                    line.runs.add(rr.copy(sb.toString()));
                    sb.setLength(0);
                }
                rr = c.run;
                sb.appendCodePoint(c.cp);
                line.width += c.w;
                if (c.cp == ' ' || c.cp == '　') line.spaces++;
                line.chars++;
            }
            if (sb.length() > 0) line.runs.add(rr.copy(sb.toString()));
            for (int i = end - 1; i >= 0 && isSpace(l.get(i).cp); i--) line.trailSpace += l.get(i).w;
            for (Run r : line.runs) {
                line.ascent = Math.max(line.ascent, r.spec.ascent * r.size);
                line.descent = Math.max(line.descent, r.spec.descent * r.size);
            }
            if (line.runs.isEmpty()) {
                // 空行の高さは、直前の行の先頭ラン (なければ行の文字またはブロック先頭のラン) から決める
                Run base = l.isEmpty() ? (runs.isEmpty() ? null : runs.get(0)) : l.get(0).run;
                if (li > 0 && !out.isEmpty() && !out.get(out.size() - 1).runs.isEmpty()) base = out.get(out.size() - 1).runs.get(0);
                if (base != null) {
                    line.blankBase = base;
                    line.ascent = base.spec.ascent * base.size;
                    line.descent = base.spec.descent * base.size;
                }
            }
            out.add(line);
        }
        return out;
    }

    /**
     * 1 行を描く。
     *
     * @param x        行の左端 (pt)
     * @param baseline ベースライン (pt)
     * @param extra    行に追加で配る幅 (pt)
     * @param mode     追加幅の配り方 (0 なし、1 空白ごと (両端揃え)、2 全文字の間 (均等割り付け))
     */
    public static void drawLine(PdfCanvas cv, Line line, double x, double baseline, float extra, int mode) throws java.io.IOException {
        double pos = x;
        int gaps = mode == 1 ? line.spaces : mode == 2 ? distGaps(line) : 0;
        float per = gaps > 0 ? extra / gaps : 0;
        int[] all = mode == 2 ? codePoints(line) : null;
        int k = 0; // 行の中の文字の位置 (均等割り付けで次の文字を見るため)
        for (Run r : line.runs) {
            float sz = r.effSize();
            double base = baseline - (r.baselineShift > 0 ? r.size * 0.33 : r.baselineShift < 0 ? -r.size * 0.14 : 0);
            double start = pos;
            if (mode == 0 || per == 0) {
                cv.text(r.text, pos, base, r.spec, sz, r.color, 1, r.spacing);
                pos += width(r, cv.fonts);
            } else {
                // 1 文字ずつ置く (最後の文字の後の隙間は呼び出し側の幅計算で扱う)
                for (int i = 0; i < r.text.length(); ) {
                    int cp = r.text.codePointAt(i);
                    i += Character.charCount(cp);
                    cv.text(new String(Character.toChars(cp)), pos, base, r.spec, sz, r.color, 1, 0);
                    pos += r.spec.charWidth(cp, sz, cv.fonts) + r.spacing;
                    if (mode == 2) {
                        if (k + 1 < all.length && distGapAfter(cp, all[k + 1])) pos += per;
                    } else if (cp == ' ' || cp == '　') {
                        pos += per;
                    }
                    k++;
                }
            }
            decorate(cv, r, start, pos, base);
        }
    }

    /** 下線と取り消し線を描く (x0〜x1、ベースライン基準)。 */
    static void decorate(PdfCanvas cv, Run r, double x0, double x1, double baseline) throws java.io.IOException {
        if (r.underline == 0 && !r.strike) return;
        float t = Math.max(0.5f, r.size * 0.06f);
        PdfCanvas.Stroke st = new PdfCanvas.Stroke(t);
        if (r.underline >= 1) {
            double y = baseline + r.size * 0.12;
            cv.line(x0, y, x1, y, st, r.color);
            if (r.underline == 2) cv.line(x0, y + t * 2, x1, y + t * 2, st, r.color);
        }
        if (r.strike) {
            double y = baseline - r.size * 0.3;
            cv.line(x0, y, x1, y, st, r.color);
        }
    }
}
