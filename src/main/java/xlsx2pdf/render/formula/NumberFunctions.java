package xlsx2pdf.render.formula;

import java.util.Locale;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 数値と文字の変換の関数 (PHI・GESTEP・DELTA・BASE・DECIMAL・ARABIC・ROMAN・UNICHAR・UNICODE)。 */
final class NumberFunctions {
    private NumberFunctions() {
    }

    // ---- 数値と文字の変換 ----
    static final double NORM_PDF_K = 0.39894228040143267793994998163520656458206656050464; // 1 / √(2π)
    static final String RADIX_DIGITS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    static final double BASE_MAX = 9007199254740991d; // 2^53 - 1 (double で整数が厳密に扱える上限)

    // ---- PHI: 標準正規分布の確率密度 (NORM.S.DIST(x,FALSE) と同じ) ----

    static ValueEval phi(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1)
            return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec);
        return new NumberEval(Math.exp(-x * x / 2) * NORM_PDF_K);
    }

    // ---- GESTEP: しきい値以上なら 1 (step を省略すると 0) ----

    static ValueEval gestep(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 1 || a.length > 2)
            return ErrorEval.VALUE_INVALID;
        // TRUE/FALSE は暗黙に数値化されない (実測: GESTEP(TRUE,0) = #VALUE!)
        if (isBool(a[0], ec) || (a.length == 2 && isBool(a[1], ec)))
            return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec);
        double step = a.length == 2 ? num(a[1], ec) : 0;
        return new NumberEval(x >= step ? 1 : 0);
    }

    /** TRUE/FALSE を渡されたら true (GESTEP・DECIMAL などの型チェック用)。 */
    static boolean isBool(ValueEval v, OperationEvaluationContext ec) throws EvaluationException {
        return scalar(v, ec) instanceof org.apache.poi.ss.formula.eval.BoolEval;
    }

    // ---- BASE / DECIMAL: n 進表記 (2〜36) ----

    static ValueEval base(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 2 || a.length > 3)
            return ErrorEval.VALUE_INVALID;
        double dn = num(a[0], ec), dr = num(a[1], ec);
        int radix = (int) Math.floor(dr);
        if (radix < 2 || radix > 36 || dn < 0 || dn > BASE_MAX)
            return ErrorEval.NUM_ERROR;
        int min = 1; // 省略は 1 桁。0 を明示すると、0 は空文字になる
        if (a.length == 3) {
            double dm = num(a[2], ec);
            if (dm < 0 || dm > 255)
                return ErrorEval.NUM_ERROR; // Excel は 255 文字を超えると #NUM! (巨大な桁数で 0 を詰め続けない)
            min = (int) Math.floor(dm);
        }
        long n = (long) Math.floor(dn);
        StringBuilder sb = new StringBuilder();
        if (n == 0 && min > 0)
            sb.append('0');
        for (long v = n; v > 0; v /= radix)
            sb.append(RADIX_DIGITS.charAt((int) (v % radix)));
        while (sb.length() < min)
            sb.append('0');
        if (sb.length() > 255)
            return ErrorEval.NUM_ERROR; // Excel は 255 文字を超えると #NUM!
        return new StringEval(sb.reverse().toString());
    }

    static ValueEval decimal(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 2 || a.length > 3)
            return ErrorEval.VALUE_INVALID;
        // 先頭の空白とタブは読み飛ばす。後ろの空白・符号・指数表記は #NUM!。空文字は 0。TRUE は #NUM!。
        // 16 進に限り "0x"・"x" の接頭辞を外す。255 文字を超えると #VALUE!、値が大きすぎると #NUM!
        if (isBool(a[0], ec))
            return ErrorEval.NUM_ERROR;
        double dr = num(a[1], ec);
        int radix = (int) Math.floor(dr);
        if (radix < 2 || radix > 36)
            return ErrorEval.NUM_ERROR;
        String t = str(a[0], ec);
        if (t.length() > 255)
            return ErrorEval.VALUE_INVALID;
        t = t.replaceFirst("^[ \\t]+", "");
        if (t.isEmpty())
            return new NumberEval(0);
        int i = 0;
        if (radix == 16) {
            if (t.length() >= 2 && t.charAt(0) == '0' && (t.charAt(1) == 'x' || t.charAt(1) == 'X'))
                i = 2;
            else if (t.charAt(0) == 'x' || t.charAt(0) == 'X')
                i = 1;
        }
        double acc = 0;
        for (; i < t.length(); i++) {
            int d = RADIX_DIGITS.indexOf(Character.toUpperCase(t.charAt(i)));
            if (d < 0 || d >= radix)
                return ErrorEval.NUM_ERROR; // 基数を超える数字・空白・記号は #NUM! (実測)
            acc = acc * radix + d;
        }
        if (Double.isNaN(acc) || Double.isInfinite(acc))
            return ErrorEval.NUM_ERROR;
        return new NumberEval(acc);
    }

    // ---- ARABIC: ローマ数字 → 数値 (ROMAN の逆) ----

    static ValueEval arabic(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1)
            return ErrorEval.VALUE_INVALID;
        String t = str(a[0], ec).trim().toUpperCase(Locale.ROOT);
        if (t.isEmpty())
            return new NumberEval(0); // 実測: ARABIC("") = 0
        boolean neg = false;
        int i = 0;
        if (t.charAt(0) == '-') { // 符号は先頭の - 1 つだけ
            neg = true;
            i = 1;
        }
        if (i >= t.length())
            return ErrorEval.NUM_ERROR; // "-" だけは #NUM!
        // 実測 (ARABIC("IIV") = 3、ARABIC("IID") = 498): 「右側に自分より大きい数字が 1 個でもあれば引く」
        // 減算表記 (IV・XC・CM) を再現しつつ、正しくない並びでも同じ値になる。
        int[] vs = new int[t.length() - i];
        for (int k = 0; k < vs.length; k++) {
            int v = romanValue(t.charAt(i + k));
            if (v < 0)
                return ErrorEval.VALUE_INVALID; // 知らない文字は #VALUE!
            vs[k] = v;
        }
        int sum = 0;
        for (int k = 0; k < vs.length; k++) {
            int mx = 0;
            for (int j = k + 1; j < vs.length; j++)
                if (vs[j] > mx)
                    mx = vs[j];
            sum += vs[k] < mx ? -vs[k] : vs[k];
        }
        return new NumberEval(neg ? -sum : sum);
    }

    static int romanValue(char c) {
        return switch (c) {
            case 'I' -> 1;
            case 'V' -> 5;
            case 'X' -> 10;
            case 'L' -> 50;
            case 'C' -> 100;
            case 'D' -> 500;
            case 'M' -> 1000;
            default -> -1;
        };
    }

    // ---- UNICHAR / UNICODE ----

    static ValueEval unichar(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1)
            return ErrorEval.VALUE_INVALID;
        double d = num(a[0], ec);
        if (Double.isNaN(d) || d < 1 || d > 0x10FFFF)
            return ErrorEval.VALUE_INVALID; // 実測: 0・0.9・1114112 は #VALUE! (1.9 は切り捨てて U+0001)
        int c = (int) Math.floor(d);
        // 実測: 面の末尾の非文字 (U+FFFE/U+FFFF, U+1FFFE…) とサロゲート域、U+FDD0〜U+FDEF は #N/A。
        //  U+F8FF (私用域) と U+FDF0 は文字として返る (63703・65008 が一致することを確認済み)。
        if ((c & 0xFFFE) == 0xFFFE || (c >= 0xD800 && c <= 0xDFFF) || (c >= 0xFDD0 && c <= 0xFDEF))
            return ErrorEval.NA;
        return new StringEval(new String(Character.toChars(c)));
    }

    static ValueEval unicode(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1)
            return ErrorEval.VALUE_INVALID;
        String t = str(a[0], ec);
        if (t.isEmpty())
            return ErrorEval.VALUE_INVALID;
        return new NumberEval(t.codePointAt(0));
    }


    static ValueEval delta(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 1 || a.length > 2)
            return ErrorEval.VALUE_INVALID;
        if (isBool(a[0], ec) || (a.length == 2 && isBool(a[1], ec)))
            return ErrorEval.VALUE_INVALID; // 論理値は #VALUE!
        double x = num(a[0], ec);
        double y = a.length >= 2 ? num(a[1], ec) : 0;
        return new NumberEval(x == y ? 1 : 0);
    }

    /** ROMAN(数値, [形式]) — POI 版は 2 引数を渡さないと #VALUE! になるため自前で実装。 */
    /**
     * ROMAN(数値, [書式])。書式は 0 (省略・TRUE) が正式、1〜3 は段階的に短く、4 (FALSE) が最も短い。
     * 数値は 0〜3999 (小数は切り捨て)。0 は空文字、負の数と 4000 以上は #VALUE!。
     */
    static ValueEval roman(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 1 || a.length > 2)
            return ErrorEval.VALUE_INVALID;
        int form = 0;
        if (a.length == 2 && !omitted(a, 1)) {
            ValueEval f = scalar(a[1], ec);
            if (f instanceof org.apache.poi.ss.formula.eval.BoolEval b)
                form = b.getBooleanValue() ? 0 : 4;
            else {
                form = (int) Math.floor(num(a[1], ec));
                if (form < 0 || form > 4)
                    return ErrorEval.VALUE_INVALID;
            }
        }
        double dn = num(a[0], ec);
        if (dn < 0 || dn >= 4000)
            return ErrorEval.VALUE_INVALID;
        return new StringEval(roman((int) Math.floor(dn), form));
    }

    private static final String ROMAN_CHARS = "MDCLXVI";
    private static final int[] ROMAN_VALUES = {1000, 500, 100, 50, 10, 5, 1};

    /** ローマ数字にする。形式 mode が大きいほど、4 や 9 の桁で引き算の記法 (IL・XM など) を広く使う。 */
    static String roman(int value, int mode) {
        StringBuilder sb = new StringBuilder();
        int val = value;
        final int last = ROMAN_VALUES.length - 1;
        for (int i = 0; i <= last / 2; i++) {
            int idx = 2 * i;
            int digit = val / ROMAN_VALUES[idx];
            if (digit % 5 == 4) {
                int big = digit == 4 ? idx - 1 : idx - 2; // 引かれる側 (500 や 1000)
                int steps = 0;
                while (steps < mode && idx < last) {
                    steps++;
                    if (ROMAN_VALUES[big] - ROMAN_VALUES[idx + 1] <= val)
                        idx++;
                    else
                        steps = mode;
                }
                sb.append(ROMAN_CHARS.charAt(idx)).append(ROMAN_CHARS.charAt(big));
                val = val + ROMAN_VALUES[idx] - ROMAN_VALUES[big];
            } else {
                if (digit > 4)
                    sb.append(ROMAN_CHARS.charAt(idx - 1));
                for (int j = 0; j < digit % 5; j++)
                    sb.append(ROMAN_CHARS.charAt(idx));
                val %= ROMAN_VALUES[idx];
            }
        }
        return sb.toString();
    }
}
