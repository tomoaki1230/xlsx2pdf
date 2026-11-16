package xlsx2pdf.render;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

/** セルと図形で共通の、書式付き文字列 (ラン) と行分割。 */
public final class TextLayout {
    /** セルの上付きの上げ幅・下付きの下げ幅 (元の大きさに対する比。Excel の印刷の実測、9〜20pt)。 */
    static final double SUP_RISE = 0.497, SUB_DROP = 0.198;
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
        /** 会計用の下線 (下線を文字から離して引く)。 */
        public boolean accounting;
        /** 下線を引く横の範囲 (NaN なら文字の幅。会計用の下線の文字列のセルは、セルの幅)。 */
        public double ulFrom = Double.NaN, ulTo = Double.NaN;
        /** 取り消し線。 */
        public boolean strike;
        /** 上付き (1)・下付き (-1)。 */
        public int baselineShift;
        /** セルの文字の上付き・下付き (Excel の実測: 大きさは 66% を整数に丸め、上付きは 0.5 倍、下付きは 0.198 倍ずらす)。 */
        public boolean cellShift;
        /** 追加の文字間隔 (pt)。 */
        public float spacing;
        /** タブのラン (text が "\t") の幅 (pt)。タブでなければ負。 */
        public float tabW = -1;
        /** 文字の影の色とずれ (pt。影がなければ null)。ぼかしは描かない。 */
        public Color shadow;
        public float shadowDx, shadowDy;

        /** ランを作る。 */
        public Run(String text, FontManager.Spec spec, float size, Color color) {
            this.text = ComplexText.cluster(text);
            this.spec = spec;
            this.size = size;
            this.color = color;
        }

        /** 書式を引き継いで、文字列だけ置き換えたランを返す。 */
        public Run copy(String t) {
            Run r = new Run(t, spec, size, color);
            r.underline = underline;
            r.accounting = accounting;
            r.ulFrom = ulFrom;
            r.ulTo = ulTo;
            r.strike = strike;
            r.baselineShift = baselineShift;
            r.cellShift = cellShift;
            r.spacing = spacing;
            r.fontName = fontName;
            r.ptSize = ptSize;
            r.tabW = tabW;
            r.shadow = shadow;
            r.shadowDx = shadowDx;
            r.shadowDy = shadowDy;
            return r;
        }

        /** 実際に描くサイズ (上付き・下付きは 66%)。 */
        float effSize() { return baselineShift == 0 ? size : cellShift ? Math.max(1, Math.round(size * 0.66f)) : size * 0.66f; }

        /** 上付き・下付きのずれ (pt、上が正)。 */
        double shift() {
            if (baselineShift == 0) return 0;
            if (cellShift) return baselineShift > 0 ? size * SUP_RISE : -size * SUB_DROP;
            return baselineShift > 0 ? size * 0.33 : -size * 0.14;
        }
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

    /**
     * タブ位置。x (文字の左端からの位置 pt) より右の最初の位置と種類 (0 左、1 中央、2 右、3 小数点) を返す。
     */
    public interface TabStops {
        /** @return {位置, 種類} */
        float[] next(float x);
    }

    /** ランの幅 (pt。文字間隔を含む)。 */
    public static float width(Run r, FontManager fm) {
        if (r.tabW >= 0) return r.tabW;
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
        return layout(runs, maxWidth, wrap, keepNewlines, fm, null, 0, 0);
    }

    /** 左揃えとみなしたタブの幅 (行分割の判断用)。 */
    private static float tabLeft(TabStops tabs, float x) {
        float[] t = tabs.next(x);
        return Math.max(0, t[0] - x);
    }

    /**
     * ランを行に分割する (タブ位置つき)。tabs が null ならタブは普通の文字として扱う。
     *
     * @param firstOffset 段落の最初の行の左端の位置 (タブ位置の基準からの pt)
     * @param otherOffset 2 行目以降の左端の位置
     */
    public static List<Line> layout(List<Run> runs, float maxWidth, boolean wrap, boolean keepNewlines, FontManager fm,
                                    TabStops tabs, float firstOffset, float otherOffset) {
        return layout(runs, maxWidth, wrap, keepNewlines, fm, tabs, firstOffset, otherOffset, false, false);
    }

    /**
     * 行数の数え方 (行の高さの数え直し用): 全角の空白を行末にぶら下げず、普通の文字と同じく次の行へ送る。送った後の行頭の
     * 空白は捨てるが、空白だけで空のまま改行に来た行は 1 行に数える。Excel の行の高さの数え直しの実測
     * (munp_eiyoukanrihoukokusyo_fukushi の「…大切である。」+ 全角の空白 15 個 + 改行で 1 行増える)。描くときはぶら下げる。
     */
    public static List<Line> layoutCounting(List<Run> runs, float maxWidth, FontManager fm) {
        return layout(runs, maxWidth, true, true, fm, null, 0, 0, WIDE_SPACE_WRAPS, true);
    }

    /** 行数の数え方で全角の空白を折り返すか (試した規則。評価の組で合わないので使わない)。 */
    static boolean WIDE_SPACE_WRAPS = false;

    private static List<Line> layout(List<Run> runs, float maxWidth, boolean wrap, boolean keepNewlines, FontManager fm,
                                     TabStops tabs, float firstOffset, float otherOffset, boolean wideSpaceWraps, boolean pixelAdvance) {
        List<Ch> chars = new ArrayList<>();
        for (Run r : runs) {
            String t = r.text;
            for (int i = 0; i < t.length(); ) {
                int cp = t.codePointAt(i);
                i += Character.charCount(cp);
                if (cp == '\r') continue;
                if (cp == '\n' && !keepNewlines) continue;
                float w = cp == '\n' ? 0 : r.spec.charWidth(cp, r.effSize(), fm) + r.spacing;
                // 画面の文字の送りは整数の画素 (大きさ pt = ppem × 0.75 のとき、送り px = 幅 pt ÷ 0.75 を丸める)
                if (pixelAdvance && w > 0) w = Math.round(w / 0.75f) * 0.75f;
                chars.add(new Ch(cp, r, w));
            }
        }
        List<List<Ch>> lines = new ArrayList<>();
        List<Boolean> paraEnd = new ArrayList<>();
        List<Ch> cur = new ArrayList<>();
        float curW = 0;
        int lastBreak = -1; // 改行できる位置 (cur の添字。行はその直前で終わる)
        boolean firstLine = true;
        boolean dropLead = false;   // wideSpaceWraps: 折り返した直後の行頭の空白を捨てる
        for (Ch c : chars) {
            if (c.cp == '\n') {
                lines.add(cur);
                paraEnd.add(true);
                cur = new ArrayList<>();
                curW = 0;
                lastBreak = -1;
                firstLine = true;
                dropLead = false;
                continue;
            }
            if (dropLead && isSpace(c.cp)) continue;
            dropLead = false;
            if (c.cp == '\t' && tabs != null) c = new Ch('\t', c.run, tabLeft(tabs, (firstLine ? firstOffset : otherOffset) + curW));
            if (wrap && maxWidth > 0 && !cur.isEmpty()) {
                int prev = cur.get(cur.size() - 1).cp;
                if (breakable(prev, c.cp)) lastBreak = cur.size();
                if (curW + c.w > maxWidth + 0.01 && (wideSpaceWraps ? c.cp != ' ' : !isSpace(c.cp))) {
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
                    firstLine = false;
                    for (int k = 0; k < cur.size(); k++) {
                        Ch x = cur.get(k);
                        if (x.cp == '\t' && tabs != null) cur.set(k, x = new Ch('\t', x.run, tabLeft(tabs, otherOffset + curW)));
                        curW += x.w;
                    }
                    lastBreak = -1;
                    if (c.cp == '\t' && tabs != null) c = new Ch('\t', c.run, tabLeft(tabs, otherOffset + curW));
                    if (wideSpaceWraps && cur.isEmpty() && isSpace(c.cp)) { dropLead = true; continue; }
                }
            }
            cur.add(c);
            curW += c.w;
        }
        lines.add(cur);
        paraEnd.add(true);

        List<Line> out = new ArrayList<>();
        boolean paraStart = true;
        for (int li = 0; li < lines.size(); li++) {
            List<Ch> l = lines.get(li);
            if (tabs != null) placeTabs(l, tabs, paraStart ? firstOffset : otherOffset);
            paraStart = paraEnd.get(li);
            Line line = new Line();
            line.endsParagraph = paraEnd.get(li);
            int end = l.size();
            if (wrap) while (end > 0 && isSpace(l.get(end - 1).cp)) end--;
            // 同じランの文字をまとめ直す
            StringBuilder sb = new StringBuilder();
            Run rr = null;
            for (int i = 0; i < end; i++) {
                Ch c = l.get(i);
                if (c.cp == '\t' && tabs != null) {
                    // タブは幅だけのラン
                    if (sb.length() > 0) { line.runs.add(rr.copy(sb.toString())); sb.setLength(0); }
                    Run t = c.run.copy("\t");
                    t.tabW = c.w;
                    line.runs.add(t);
                    line.width += c.w;
                    line.chars++;
                    rr = null;
                    continue;
                }
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
                if (r.tabW >= 0) continue;
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
     * 行のタブの幅を、タブ位置の種類で決め直す。中央・右・小数点は、次のタブ (または行末) までの文字列を位置に合わせる
     * (小数点は最初の . の前まで)。届かなければ幅 0。
     */
    private static void placeTabs(List<Ch> l, TabStops tabs, float offset) {
        float x = offset;
        for (int i = 0; i < l.size(); i++) {
            Ch c = l.get(i);
            if (c.cp != '\t') { x += c.w; continue; }
            float[] t = tabs.next(x);
            float seg = 0, dec = -1;
            for (int j = i + 1; j < l.size() && l.get(j).cp != '\t'; j++) {
                if (dec < 0 && l.get(j).cp == '.') dec = seg;
                seg += l.get(j).w;
            }
            float w = switch ((int) t[1]) {
                case 1 -> t[0] - seg / 2 - x;
                case 2 -> t[0] - seg - x;
                case 3 -> t[0] - (dec < 0 ? seg : dec) - x;
                default -> t[0] - x;
            };
            w = Math.max(0, w);
            l.set(i, new Ch('\t', c.run, w));
            x += w;
        }
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
            double base = baseline - r.shift();
            double start = pos;
            if (r.tabW >= 0) {
                pos += r.tabW;
                continue;
            }
            if (r.shadow != null) cv.text(r.text, pos + r.shadowDx, base + r.shadowDy, r.spec, sz, r.shadow, 1, r.spacing);
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
            // 会計用の下線は、普通の下線より大きさ × 0.207 下 (Excel の印刷: 11pt で 2.28pt)
            // 会計用の下線の下げ幅は書体のディセントに比例 (游ゴシックのディセント 0.445em で 0.207em。Excel の印刷 exp96)
            double acc = r.spec != null && r.spec.descent > 0 ? r.size * r.spec.descent * 0.465 : r.size * 0.207;
            double y = baseline + r.size * 0.12 + (r.accounting ? acc : 0);
            double u0 = Double.isNaN(r.ulFrom) ? x0 : r.ulFrom, u1 = Double.isNaN(r.ulTo) ? x1 : r.ulTo;
            cv.line(u0, y, u1, y, st, r.color);
            if (r.underline == 2) cv.line(u0, y + t * 2, u1, y + t * 2, st, r.color);
        }
        if (r.strike) {
            double y = baseline - r.size * 0.3;
            cv.line(x0, y, x1, y, st, r.color);
        }
    }
}
