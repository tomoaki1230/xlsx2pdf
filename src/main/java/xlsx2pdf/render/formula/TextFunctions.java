package xlsx2pdf.render.formula;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.functions.Function;
import xlsx2pdf.render.ExcelFormat;
import xlsx2pdf.render.FontManager;
import xlsx2pdf.render.RenderContext;
import xlsx2pdf.render.ValueFormatter;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.DateFunctions.*;

/** TEXT・TEXTBEFORE・TEXTAFTER。 */
final class TextFunctions {
    private TextFunctions() {
    }

    private static boolean installed;

    /** POI の TEXT を、PDF の描画と同じ Excel の表示形式の処理 (ExcelFormat) で置き換える。 */
    /** POI の TEXT の置き換え (JVM 全体で一度だけ)。 */
    static void install() {
        if (installed) return;
        installed = true;
        try {
            var meta = org.apache.poi.ss.formula.function.FunctionMetadataRegistry.getFunctionByName("TEXT");
            var field = FunctionEval.class.getDeclaredField("functions");
            field.setAccessible(true);
            Function[] table = (Function[]) field.get(null);
            Function orig = table[meta.getIndex()];
            table[meta.getIndex()] = (args, row, col) -> {
                try {
                    return text(args, row, col);
                } catch (EvaluationException e) {
                    return e.getErrorEval();
                } catch (RuntimeException e) {
                    return orig.evaluate(args, row, col);
                }
            };
        } catch (ReflectiveOperationException | RuntimeException e) {
            // 置き換えられなければ POI の TEXT のまま
        }
    }

    static ValueEval text(ValueEval[] a, int row, int col) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        ValueEval v = OperandResolver.getSingleValue(a[0], row, col);
        String fmt = OperandResolver.coerceValueToString(OperandResolver.getSingleValue(a[1], row, col));
        if (v instanceof BoolEval b) return new StringEval(b.getBooleanValue() ? "TRUE" : "FALSE");
        if (fmt.isEmpty()) return new StringEval("");
        // 日本語版の Excel の TEXT は、日本語の書式の書き方だけを受け付ける (General・[Red] などの英語は #VALUE!。標準は G/標準、色は [赤])。
        // 日本語版で作ったブック (標準のスタイルの名前が「標準」) だけにこの規則を使う。それ以外のブックは英語の書き方も受け付ける
        if (FormulaContext.workbook() != null && RenderContext.normalStyleJa(FormulaContext.workbook())) {
            fmt = jaFormat(fmt);
            if (fmt == null) return ErrorEval.VALUE_INVALID;
        }
        Double d;
        String str = null;
        if (v instanceof StringEval se) {
            str = se.getStringValue();
            d = numberOrDate(str);
        } else {
            d = v == BlankEval.instance || v == org.apache.poi.ss.formula.eval.MissingArgEval.instance ? 0 : OperandResolver.coerceValueToDouble(v);
        }
        // 「@」(と @ だけを含む書式) は数値を標準の形で返し、文字列はそのまま返す
        boolean at = fmt.trim().equals("@");
        if (d != null && (at || ValueFormatter.textOnly(fmt))) return new StringEval(ValueFormatter.general(d, 15));
        if (d == null && at) return new StringEval(str);
        // [h]・[m]・[s] だけの書式は、負の値に - を付けて経過時間を返す
        if (d != null && d < 0 && fmt.trim().matches("(?i)\\[(h+|m+|s+)\\]")) {
            ExcelFormat.Out neg = new ExcelFormat(FormulaContext.is1904()).format(-d, fmt);
            return new StringEval("-" + neg.text);
        }
        ExcelFormat ef = new ExcelFormat(FormulaContext.is1904());
        ExcelFormat.Out o = d != null ? ef.format(d, fmt) : ef.formatText(str, fmt);
        if (o == null) return new StringEval(str);
        // 日付にできない値 (負・9999/12/31 を超える) は、セルでは # で埋まるが、TEXT は #VALUE!
        if (o.text.equals("#".repeat(8))) return ErrorEval.VALUE_INVALID;
        StringBuilder sb = new StringBuilder();
        for (char c : o.text.toCharArray()) {
            if (c == ExcelFormat.FILL) continue;                       // 繰り返し (*x) は TEXT では何も出さない
            sb.append(c >= FontManager.PAD_BASE && c < FontManager.PAD_BASE + 0x100 || c == '\u2007' || c == '\u2009' ? ' ' : c);   // _x は空白 1 つ
        }
        String r = sb.toString();
        // 分数 (# ?/?) は整数部がなくても、整数部と分子の間の空白が残る (0.5 → " 1/2")。セルの表示ではこの空白を出さない
        if (r.matches("-?\\d+/\\s*\\d+") && fmt.matches(".*#\\s+\\?.*/.*")) r = r.startsWith("-") ? "- " + r.substring(1) : " " + r;
        return new StringEval(r);
    }

    /** 日本語版の TEXT の書式を ExcelFormat の書き方にする (英語の書き方や使えない書き方は null)。 */
    static String jaFormat(String fmt) {
        String plain = fmt.replaceAll("\"[^\"]*\"", "");
        if (plain.matches("(?is).*general.*") || plain.contains("午前/午後")) return null;
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("\\[([^\\]]*)\\]").matcher(plain);
        while (m.find()) {
            String t = m.group(1).toLowerCase(Locale.ROOT);
            if (t.matches("red|blue|green|cyan|magenta|yellow|white|black|color\\d+")) return null;
        }
        return fmt.replaceAll("\\[(赤|青|緑|水|黄|白|黒|紫|色\\d+)\\]", "").replace("G/標準", "General");
    }

    /** 文字列を数値に読む (数値・日付の形の文字列。読めなければ null)。 */
    static Double numberOrDate(String s) {
        String t = s.trim();
        if (t.isEmpty()) return null;
        try {
            return Double.valueOf(t);
        } catch (NumberFormatException e) {
            // 日付かもしれない
        }
        try {
            return dateArg(new StringEval(t));
        } catch (EvaluationException e) {
            return null;
        }
    }


    // ---- TEXTBEFORE / TEXTAFTER ----

    static ValueEval textSplit(ValueEval[] a, OperationEvaluationContext ec, boolean before) throws EvaluationException {
        if (a.length < 2 || a.length > 6) return ErrorEval.VALUE_INVALID;
        String text = str(a[0], ec);
        List<String> delims = new ArrayList<>();
        ValueEval d = a[1];
        if (d instanceof AreaEval ar) {
            for (int r = 0; r < ar.getHeight(); r++)
                for (int c = 0; c < ar.getWidth(); c++) delims.add(OperandResolver.coerceValueToString(ar.getRelativeValue(r, c)));
        } else {
            delims.add(str(d, ec));
        }
        int inst = omitted(a, 2) ? 1 : (int) num(a[2], ec);
        boolean ci = !omitted(a, 3) && num(a[3], ec) != 0;
        boolean matchEnd = !omitted(a, 4) && num(a[4], ec) != 0;
        if (text.isEmpty() && matchEnd && delims.stream().anyMatch(x -> !x.isEmpty())) throw new EvaluationException(ErrorEval.VALUE_INVALID);
        if (inst == 0) throw new EvaluationException(ErrorEval.VALUE_INVALID);
        // 出現回数を書いたとき、文字列に入りきらない回数 (長さ − 区切りの長さ + 1 を超える) は #VALUE! (Excel の結果)
        int minLen = delims.stream().mapToInt(String::length).filter(n -> n > 0).min().orElse(0);
        if (!omitted(a, 2) && minLen > 0 && Math.abs(inst) > Math.max(0, text.length() - minLen + 1))
            throw new EvaluationException(ErrorEval.VALUE_INVALID);
        String hay = ci ? text.toLowerCase(Locale.ROOT) : text;
        List<String> ds = new ArrayList<>();
        for (String s : delims) ds.add(ci ? s.toLowerCase(Locale.ROOT) : s);
        boolean allEmpty = ds.stream().allMatch(String::isEmpty);
        int start = -1, end = -1;
        if (allEmpty) {
            // 区切りが空: 先頭 (出現回数が正) か末尾 (負) に一致する。1 文字の文字列は正でも末尾 (Excel の結果)
            start = end = inst > 0 && text.length() != 1 ? 0 : text.length();
        } else {
            int n = Math.abs(inst), found = 0;
            if (inst > 0) {
                int pos = 0;
                while (pos <= hay.length()) {
                    int bp = -1, bl = 0;
                    for (String s : ds) {
                        if (s.isEmpty()) continue;
                        int p = hay.indexOf(s, pos);
                        if (p >= 0 && (bp < 0 || p < bp || p == bp && s.length() > bl)) {
                            bp = p;
                            bl = s.length();
                        }
                    }
                    if (bp < 0) break;
                    if (++found == n) {
                        start = bp;
                        end = bp + bl;
                        break;
                    }
                    pos = bp + bl;
                }
                if (start < 0 && matchEnd && found + 1 == n) start = end = text.length();
            } else {
                int limit = hay.length();
                while (limit >= 0) {
                    int bp = -1, bl = 0;
                    for (String s : ds) {
                        if (s.isEmpty()) continue;
                        int p = hay.lastIndexOf(s, limit - s.length());
                        if (p >= 0 && p + s.length() <= limit && (bp < 0 || p + s.length() > bp + bl || p + s.length() == bp + bl && s.length() > bl)) {
                            bp = p;
                            bl = s.length();
                        }
                    }
                    if (bp < 0) break;
                    if (++found == n) {
                        start = bp;
                        end = bp + bl;
                        break;
                    }
                    limit = bp;
                }
                if (start < 0 && matchEnd && found + 1 == n) start = end = 0;
            }
        }
        if (start < 0) {
            if (!omitted(a, 5)) return scalar(a[5], ec);
            return ErrorEval.NA;
        }
        return new StringEval(before ? text.substring(0, start) : text.substring(end));
    }
}
