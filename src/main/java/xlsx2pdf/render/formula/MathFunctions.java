package xlsx2pdf.render.formula;

import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 数学の関数 (ERF 系・GAMMA 系・PERMUT 系・BIT 系)。 */
final class MathFunctions {
    private MathFunctions() {
    }

    // ---- ERF・ERFC・GAMMA・GAMMALN ----

    static final double K_2_OVER_SQRT_PI = 1.1283791670955125738961589031215451716536964511283; // 2 / √π

    /** mode 0: ERF(x)・ERF(下限,上限)、1: ERF.PRECISE、2: ERFC・ERFC.PRECISE。 */
    static ValueEval erfFn(ValueEval[] a, OperationEvaluationContext ec, int mode) throws EvaluationException {
        for (ValueEval x : a) if (x instanceof org.apache.poi.ss.formula.eval.BoolEval) return ErrorEval.VALUE_INVALID; // 論理値は #VALUE!
        if (mode == 2) {
            if (a.length != 1) return ErrorEval.VALUE_INVALID;
            return new NumberEval(erfc(num(a[0], ec)));
        }
        if (mode == 1) {
            if (a.length != 1) return ErrorEval.VALUE_INVALID;
            return new NumberEval(erf(num(a[0], ec)));
        }
        if (a.length == 1) return new NumberEval(erf(num(a[0], ec)));
        if (a.length == 2) {
            double lo = num(a[0], ec), up = num(a[1], ec);
            return new NumberEval(erf(up) - erf(lo));
        }
        return ErrorEval.VALUE_INVALID;
    }

    /** 誤差関数。|x| ≦ 2.5 は冪級数、それより大きいときは 1 - erfc (Excel と同じ 15 桁を級数と連分数で確認)。 */
    static double erf(double x) {
        if (x < 0) return -erf(-x);
        if (x <= 2.5) return erfSeries(x);
        double c = erfcCf(x);
        return c == 0 || !Double.isFinite(c) ? 1 : 1 - c;
    }

    /** 相補誤差関数。x < 1.8 は 1 - 級数、それ以降は連分数 (小さい値を 1 から引かない)。 */
    static double erfc(double x) {
        if (x < 0) return 2 - erfc(-x);
        if (x < 1.8) return 1 - erfSeries(x);
        return erfcCf(x);
    }

    /** erf(x) = (2/√π) Σ (-1)^n x^(2n+1) / (n! (2n+1))。 */
    static double erfSeries(double x) {
        double s = x, t = x;
        for (int n = 1; n < 500; n++) {
            t *= -x * x / n;
            double term = t / (2 * n + 1);
            s += term;
            if (Math.abs(term) < 1e-19 * Math.abs(s)) break;
        }
        return K_2_OVER_SQRT_PI * s;
    }

    /** erfc(x) = (2/√π) e^(-x²) / h, h = 2x + 2/(2x + 4/(2x + 6/(…))) を下から評価した連分数。 */
    static double erfcCf(double x) {
        double a = 2 * x, h = a;
        for (int k = 200; k >= 1; k--) {
            h = a + (2.0 * k) / h;
            if (Math.abs(h) < 1e-300) h = 1e-300;
        }
        double e = Math.exp(-x * x);
        return e == 0 ? 0 : K_2_OVER_SQRT_PI * e / h;
    }

    static ValueEval gammaFn(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec);
        if (x == 0 || (x < 0 && x == Math.floor(x))) return ErrorEval.NUM_ERROR; // 0・負の整数は #NUM! (Excel)
        if (x < 0) {
            double r = Math.PI / (Math.sin(Math.PI * x) * gammaPos(1 - x)); // 反射公式
            return Double.isFinite(r) ? new NumberEval(r) : ErrorEval.NUM_ERROR;
        }
        double r = gammaPos(x);
        return Double.isFinite(r) ? new NumberEval(r) : ErrorEval.NUM_ERROR; // GAMMA(180) は溢れて #NUM! (実 Excel で確認)
    }

    static ValueEval gammaLn(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec);
        if (x <= 0) return ErrorEval.NUM_ERROR; // 正の数以外は #NUM! (Excel)
        return new NumberEval(logGamma(x));
    }

    static double gammaPos(double x) {
        return Math.exp(logGamma(x));
    }

    /** スタリングの級数で log Γ(x) (x を 12 以上へ移してから、Γ(x+1)=xΓ(x) で戻す)。 */
    static double logGamma(double x) {
        double acc = 0;
        while (x < 12) {
            acc -= Math.log(x);
            x += 1;
        }
        double iv = 1 / (x * x);
        double p = 1 / 12.0 + iv * (-1 / 360.0 + iv * (1 / 1260.0 + iv * (-1 / 1680.0
                + iv * (1 / 1188.0 + iv * (-691 / 360360.0 + iv * (1 / 156.0 + iv * (-3617 / 122400.0)))))));
        return acc + (x - 0.5) * Math.log(x) - x + 0.91893853320467274178032973640564696289183503507084 + p / x;
    }

    // ---- PERMUT・PERMUTATIONA・BIT 系・ISPMT ----

    static long trunc(double v) {
        return (long) v; // Excel と同じ切り捨て (負の小数は 0 へ寄せる)
    }

    static ValueEval permut(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        long n = trunc(num(a[0], ec)), k = trunc(num(a[1], ec));
        if (n < 0 || k < 0) return ErrorEval.NUM_ERROR;
        if (k > n) return ErrorEval.NUM_ERROR; // 取り出す数が総数より大きい (PERMUT(5,6) = #NUM! と実 Excel で確認)
        double r = 1;
        for (long i = n - k + 1; i <= n && !Double.isInfinite(r); i++) r *= i;   // 溢れたら止める (k が極端に大きいと終わらない)
        return Double.isInfinite(r) ? ErrorEval.NUM_ERROR : new NumberEval(r);
    }

    static ValueEval permuta(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 1 || a.length > 2) return ErrorEval.VALUE_INVALID;
        long n = trunc(num(a[0], ec)), k = omitted(a, 1) ? n : trunc(num(a[1], ec));
        if (n < 0 || k < 0) return ErrorEval.NUM_ERROR;
        double r = Math.pow(n, k); // 重複を許す順列は n^k
        return Double.isInfinite(r) ? ErrorEval.NUM_ERROR : new NumberEval(r);
    }

    /** BITAND・BITOR・BITXOR。小数は切り捨て、負の数 (#NUM! は Excel)。上限 (2^10 という説) は Excel で未確認なので設けない。 */
    /**
     * BITAND・BITOR・BITXOR。実 Excel は小数の引数を切り捨てずに #NUM! にする
     * (BITAND(2.7,1.2)・BITOR(5.9,2)・BITXOR(2.9,1.1) で確認)。負の数も #NUM!。
     */
    /** BIT* 関数の引数と結果の上限 (2^48 - 1)。これを超えると #NUM! (Excel 2024 で確認)。 */
    static final double BIT_VALUE_MAX = 281474976710655d;

    static ValueEval bitwise(ValueEval[] a, OperationEvaluationContext ec, int op) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double xd = num(a[0], ec), yd = num(a[1], ec);
        if (xd != Math.floor(xd) || yd != Math.floor(yd)) return ErrorEval.NUM_ERROR; // 小数は #NUM! (実 Excel)
        // 2^48 - 1 を超える値は #NUM!
        if (xd > BIT_VALUE_MAX || yd > BIT_VALUE_MAX) return ErrorEval.NUM_ERROR;
        if (xd > 9.2e18 || yd > 9.2e18) return ErrorEval.NUM_ERROR; // long に載らない
        long x = (long) xd, y = (long) yd;
        if (x < 0 || y < 0) return ErrorEval.NUM_ERROR;
        long r = op == 0 ? x & y : op == 1 ? x | y : x ^ y;
        return new NumberEval(r);
    }

    /**
     * BITLSHIFT・BITRSHIFT(数値, シフト数)。数値は 0〜2^48-1 の整数 (小数・負の数・上限超えは #NUM!)。
     * シフト数は切り捨てる。BITRSHIFT の負のシフト数は左へのシフト。Excel の結果に合わせた次の癖がある:
     * 54 個以上は左右どちらでも #NUM!、左に 49〜53 個は 0、左に 48 個までは結果が上限を超えると #NUM!。
     */
    static ValueEval bitShift(ValueEval[] a, OperationEvaluationContext ec, boolean left) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double xd = num(a[0], ec), sd = num(a[1], ec);
        if (xd != Math.floor(xd)) return ErrorEval.NUM_ERROR;
        if (xd > BIT_VALUE_MAX) return ErrorEval.NUM_ERROR;
        if (xd > 9.2e18) return ErrorEval.NUM_ERROR;
        long x = (long) xd;
        if (x < 0) return ErrorEval.NUM_ERROR;
        long s = (long) sd;
        boolean toLeft = left ? s > 0 : s < 0; // BITRSHIFT は負の個数だと左へ動かす (Excel)
        long n = Math.abs(s);
        if (n == 0 || x == 0) return new NumberEval((double) x);
        if (n >= 54) return ErrorEval.NUM_ERROR;
        if (toLeft) {
            if (n >= 49) return new NumberEval(0.0);
            if (x * Math.pow(2, n) > BIT_VALUE_MAX) return ErrorEval.NUM_ERROR; // 結果が上限を超える
            double r = x * Math.pow(2, n);
            if (Double.isInfinite(r) || Double.isNaN(r)) return ErrorEval.NUM_ERROR;
            return new NumberEval(r);
        }
        double r = x * Math.pow(2, -n); // 右へ捨てた分は切り捨て
        if (Double.isInfinite(r) || Double.isNaN(r)) return ErrorEval.NUM_ERROR;
        return new NumberEval(Math.floor(r));
    }


}
