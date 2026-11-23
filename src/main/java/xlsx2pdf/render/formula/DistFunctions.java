package xlsx2pdf.render.formula;

import java.util.List;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.SpecialMath.*;
import static xlsx2pdf.render.formula.MathFunctions.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 確率分布の関数 (*.DIST・*.INV とその旧名、検定、信頼区間)。 */
final class DistFunctions {
    private DistFunctions() {
    }

    // ---- BETA.DIST / BETADIST / BETA.INV / BETAINV ----

    /**
     * BETA.DIST(x, α, β, 累積, [A], [B]) と旧名 BETADIST(x, α, β, [A], [B]) (累積分布だけ)。
     */
    static ValueEval betaDist(ValueEval[] a, OperationEvaluationContext ec, boolean legacy) throws EvaluationException {
        if (a.length < (legacy ? 3 : 4) || a.length > (legacy ? 5 : 6)) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), al = num(a[1], ec), be = num(a[2], ec);
        int first = legacy ? 3 : 4; // A の引数の位置
        double cum = legacy ? 1 : num(a[3], ec);
        double A = omitted(a, first) ? 0 : num(a[first], ec), B = omitted(a, first + 1) ? 1 : num(a[first + 1], ec);
        if (al <= 0 || be <= 0 || A >= B) return ErrorEval.NUM_ERROR; // α・β が 0 以下、A ≧ B
        double t = (x - A) / (B - A);
        if (t < 0 || t > 1) return ErrorEval.NUM_ERROR; // x が [A,B] の外は #NUM! (Excel の文書。0/1 を返す実装もある)
        boolean c = cum != 0;
        if (!c) { // 確率密度
            double den = Math.exp(logBetaAB(al, be)) * (B - A);
            double pdf;
            if (t == 0) pdf = al <= 1 ? Double.POSITIVE_INFINITY : 0; // α = 1 でも #NUM! (Excel)
            else if (t == 1) pdf = be <= 1 ? Double.POSITIVE_INFINITY : 0; // β = 1 でも #NUM! (Excel)
            else pdf = Math.exp((al - 1) * Math.log(t) + (be - 1) * Math.log1p(-t)) / den;
            if (!Double.isFinite(pdf)) return ErrorEval.NUM_ERROR; // x=A・x=B で ∞ になる場合
            return new NumberEval(pdf);
        }
        return new NumberEval(t == 0 ? 0 : t == 1 ? 1 : betI(t, al, be));
    }

    static ValueEval betaInv(ValueEval[] a, OperationEvaluationContext ec, boolean legacy) throws EvaluationException {
        int want = legacy ? 3 : 5;
        if (a.length < want - 2 || a.length > want) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), al = num(a[1], ec), be = num(a[2], ec);
        double A = omitted(a, 3) ? 0 : num(a[3], ec), B = omitted(a, 4) ? 1 : num(a[4], ec);
        if (p <= 0 || p >= 1 || al <= 0 || be <= 0 || A >= B) return ErrorEval.NUM_ERROR; // 確率 0 と 1 は #NUM! (Excel で実測)
        double t = invert(u -> betI(u, al, be), u -> betI(1 - u, be, al), p, 0, 1);
        return new NumberEval(A + t * (B - A));
    }

    // ---- BINOM.DIST / BINOMDIST / BINOM.DIST.RANGE / CRITBINOM ----

    static ValueEval binomDist(ValueEval[] a, OperationEvaluationContext ec, boolean cumulative) throws EvaluationException {
        if (a.length < 4 || a.length > 5) return ErrorEval.VALUE_INVALID;
        int k = (int) trunc(num(a[0], ec)), n = (int) trunc(num(a[1], ec));
        double p = num(a[2], ec), cum = num(a[3], ec);
        if (k < 0 || n < 0 || p < 0 || p > 1) return ErrorEval.NUM_ERROR;
        if (k > n) return ErrorEval.NUM_ERROR; // 成功回数が試行回数を超えると #NUM! (BINOM.DIST(11,10,0.5,TRUE) と実 Excel で確認)
        if (cum == 0) return new NumberEval(binomPmf(k, n, p)); // 確率質量
        if (k == n) return new NumberEval(1);
        // 分布関数 = I_{1-p}(n-k, k+1) (二項和の閉じた形。足し算より桁が落ちない)
        return new NumberEval(betI(1 - p, n - k, k + 1d));
    }

    static ValueEval binomRange(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 3 || a.length > 4) return ErrorEval.VALUE_INVALID;
        int n = (int) trunc(num(a[0], ec));
        double p = num(a[1], ec);
        int k1 = (int) trunc(num(a[2], ec));
        int k2 = omitted(a, 3) ? k1 : (int) trunc(num(a[3], ec));
        if (n < 0 || p < 0 || p > 1 || k1 < 0 || k2 < 0 || k1 > k2 || k2 > n) return ErrorEval.NUM_ERROR; // 回数が試行回数を超えると #NUM! (Excel で実測)
        checkSteps((double) k2 - k1);
        double s = 0;
        for (int i = k1; i <= k2; i++) s += binomPmf(i, n, p);
        return new NumberEval(s);
    }

    static ValueEval critBinom(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        int n = (int) trunc(num(a[0], ec));
        double p = num(a[1], ec), al = num(a[2], ec);
        if (n < 0 || p <= 0 || p >= 1 || al <= 0 || al >= 1) return ErrorEval.NUM_ERROR;
        // 累積確率 (k について増える) が al 以上になる最小の k を二分法で探す (1 つずつ調べると n が大きいとき終わらない)
        int lo = 0, hi = n;
        while (lo < hi) {
            int k = lo + (hi - lo) / 2;
            double cdf = betI(1 - p, n - k, k + 1d);
            if (cdf >= al) hi = k; else lo = k + 1;
        }
        return new NumberEval(lo);
    }

    // ---- CHISQ.DIST / CHIDIST / CHISQ.DIST.RT / CHISQ.INV / CHISQ.INV.RT / CHITEST / CHISQ.TEST ----

    static ValueEval chi2Dist(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 3 : a.length != 2) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), df = trunc(num(a[1], ec));
        if (x < 0 || df < 1) return ErrorEval.NUM_ERROR;
        if (!hasCumulative) return new NumberEval(chiRt(x, df)); // CHIDIST は常に上側 (旧名)
        double cum = num(a[2], ec);
        if (cum == 0) { // 確率密度。x = 0 は自由度 1 で #NUM!、2 で 0.5、3 以上で 0 (Excel)
            if (x == 0 && df < 2) return ErrorEval.NUM_ERROR;
            return new NumberEval(chiPdf(x, df));
        }
        return new NumberEval(chiCdf(x, df));
    }

    static ValueEval chi2Inv(ValueEval[] a, OperationEvaluationContext ec, boolean rightTail) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), df = trunc(num(a[1], ec));
        if (p < 0 || p > 1 || df < 1) return ErrorEval.NUM_ERROR;
        if (rightTail && p == 0) return ErrorEval.NUM_ERROR; // 上側確率 0 は ∞ (Excel で #NUM! を確認)
        if (rightTail && p == 1) return new NumberEval(0); // 上側確率 1 は 0 (CHISQ.INV.RT(1,5) = 0 を Excel で確認)
        if (!rightTail && p == 1) return ErrorEval.NUM_ERROR; // 下側確率 1 (CHISQ.INV(1,5)) は #NUM! (Excel で確認)
        double v = rightTail ? chiInvRt(p, df) : chiInv(p, df);
        return new NumberEval(v);
    }

    static ValueEval chi2Test(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double[][] o = matrix(a[0]), e = matrix(a[1]);
        if (o == null || e == null) return ErrorEval.VALUE_INVALID;
        if (o.length != e.length || o[0].length != e[0].length) return ErrorEval.NA; // 大きさが違うと #N/A
        double stat = 0;
        int kept = 0;
        for (int r = 0; r < o.length; r++)
            for (int c = 0; c < o[0].length; c++) {
                if (Double.isNaN(o[r][c]) || Double.isNaN(e[r][c]))
                    continue; // 文字列・空のマスは組ごと無視 (Excel: CHISQ.TEST(A125:F125,A126:F126) = 1)
                if (e[r][c] == 0) return ErrorEval.DIV_ZERO; // 期待値が 0 (数値) のマスは #DIV/0! (Excel で確認)
                double d = o[r][c] - e[r][c];
                stat += d * d / e[r][c];
                kept++;
            }
        if (kept == 0) return ErrorEval.DIV_ZERO;
        int df = o.length > 1 && o[0].length > 1 ? (o.length - 1) * (o[0].length - 1) : o.length * o[0].length - 1;
        if (df <= 0) return ErrorEval.DIV_ZERO;
        return new NumberEval(chiRt(stat, df));
    }

    /** 2 次元の範囲 (CHITEST・FTEST 用)。範囲以外 (数値など) は null。 */
    static double[][] matrix(ValueEval v) throws EvaluationException {
        if (v instanceof RefEval re) v = re.getInnerValueEval(re.getFirstSheetIndex());
        if (!(v instanceof AreaEval ar)) return null;
        double[][] out = new double[ar.getHeight()][ar.getWidth()];
        for (int r = 0; r < ar.getHeight(); r++)
            for (int c = 0; c < ar.getWidth(); c++) {
                ValueEval x = ar.getRelativeValue(r, c);
                if (x instanceof NumberEval n) out[r][c] = n.getNumberValue();
                else if (x instanceof BoolEval b) out[r][c] = b.getBooleanValue() ? 1 : 0;
                else if (x instanceof ErrorEval e) throw new EvaluationException(e);
                else out[r][c] = Double.NaN; // 空・文字列は「使えないマス」の印 (CHISQ.TEST では組ごと無視する)
            }
        return out;
    }

    // ---- F.DIST / FDIST / F.DIST.RT / F.INV / FINV / F.INV.RT / FTEST / F.TEST ----

    static ValueEval fDist(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 4 : a.length != 3) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), d1 = trunc(num(a[1], ec)), d2 = trunc(num(a[2], ec));
        if (x < 0 || d1 < 1 || d2 < 1) return ErrorEval.NUM_ERROR;
        if (!hasCumulative) return new NumberEval(1 - fCdf(x, d1, d2)); // FDIST は常に上側 (旧名)
        double cum = num(a[3], ec);
        if (cum == 0) { // 確率密度。x = 0 は d1 = 1 で #NUM!、2 で 1、3 以上で 0 (Excel)
            if (x == 0) return d1 < 2 ? ErrorEval.NUM_ERROR : new NumberEval(d1 == 2 ? 1 : 0);
            return new NumberEval(fPdf(x, d1, d2));
        }
        return new NumberEval(fCdf(x, d1, d2));
    }

    static ValueEval fInvFn(ValueEval[] a, OperationEvaluationContext ec, boolean rightTail) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), d1 = trunc(num(a[1], ec)), d2 = trunc(num(a[2], ec));
        if (p < 0 || d1 < 1 || d2 < 1) return ErrorEval.NUM_ERROR;
        if (!rightTail && p >= 1) return ErrorEval.NUM_ERROR; // 下側確率 1 は #NUM! (F.INV(1,3,4) と実 Excel で確認)
        if (rightTail && (p < 1e-300 || p > 1)) return ErrorEval.NUM_ERROR;
        if (rightTail && p == 1) return new NumberEval(0); // 上側確率 1 = 0 (F.INV.RT(1,3,4) = 0 と実 Excel で確認)
        if (rightTail && p == 0) return ErrorEval.NUM_ERROR; // 上側確率 0 は ∞ (#NUM!)
        double v = rightTail ? fInvRt(p, d1, d2) : fInv(p, d1, d2);
        return new NumberEval(v);
    }

    /** F 検定 (FTEST と F.TEST は Excel 2010 以降同じ値)。 */
    static ValueEval fTest(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        List<Double> v1 = numbers(a[0]), v2 = numbers(a[1]);
        int n1 = v1.size(), n2 = v2.size();
        if (n1 < 2 || n2 < 2) return ErrorEval.DIV_ZERO;
        double s1 = variance(v1, n1), s2 = variance(v2, n2);
        if (s1 == 0 || s2 == 0) return ErrorEval.DIV_ZERO; // どちらの分散が 0 でも #DIV/0! (Excel で実測)
        double r = s1 / s2;
        double left = fCdf(r, n1 - 1, n2 - 1);
        return new NumberEval(Math.min(1, 2 * Math.min(left, 1 - left)));
    }

    /** 標本分散 (n-1 で割る)。 */
    static double variance(List<Double> v, int n) {
        double m = 0;
        for (double d : v) m += d;
        m /= n;
        double s = 0;
        for (double d : v) s += (d - m) * (d - m);
        return s / (n - 1);
    }

    // ---- GAMMA.DIST / GAMMADIST / GAMMA.INV / GAMMAINV ----

    static ValueEval gammaDist(ValueEval[] a, OperationEvaluationContext ec, boolean cumulative) throws EvaluationException {
        if (a.length < 4 || a.length > 5) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), al = num(a[1], ec), be = num(a[2], ec), cum = num(a[3], ec);
        if (x < 0 || al <= 0 || be <= 0) return ErrorEval.NUM_ERROR;
        double z = x / be; // 目盛ベータで割ったものが標準ガンマの引数
        if (cum == 0) { // 確率密度
            if (x == 0) return al <= 1 ? ErrorEval.NUM_ERROR : new NumberEval(0); // α が 1 以下は #NUM! (Excel)
            double pdf = Math.exp((al - 1) * Math.log(z) - z - logGamma(al)) / be;
            if (!Double.isFinite(pdf)) return ErrorEval.NUM_ERROR;
            return new NumberEval(pdf);
        }
        return new NumberEval(z == 0 ? 0 : gamP(al, z));
    }

    static ValueEval gammaInv(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), al = num(a[1], ec), be = num(a[2], ec);
        if (p < 0 || p >= 1 || al <= 0 || be <= 0) return ErrorEval.NUM_ERROR; // 確率 1 は #NUM! (GAMMA.INV(1,2,1) を Excel で確認)
        if (p == 0) return new NumberEval(0);
        double z = invert(u -> gamP(al, u), u -> gamQ(al, u), p, 0, Math.max(2, al + 4));
        return new NumberEval(z * be);
    }

    // ---- LOGNORM.DIST / LOGNORMDIST / LOGNORM.INV / LOGINV ----

    static ValueEval logNormDist(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 4 : a.length != 3) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), mu = num(a[1], ec), sd = num(a[2], ec);
        double cum = hasCumulative ? num(a[3], ec) : 1;
        if (x <= 0 || sd <= 0) return ErrorEval.NUM_ERROR;
        double z = (Math.log(x) - mu) / sd;
        if (!hasCumulative) return new NumberEval(cum == 0 ? 1 - normCdf(z) : normCdf(z)); // 旧名 LOGNORMDIST は分布関数だけ
        if (cum == 0) return new NumberEval(Math.exp(-z * z / 2) / (x * sd * Math.sqrt(2 * Math.PI)));
        return new NumberEval(normCdf(z));
    }

    static ValueEval logNormInv(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), mu = num(a[1], ec), sd = num(a[2], ec);
        if (p <= 0 || p >= 1 || sd <= 0) return ErrorEval.NUM_ERROR;
        return new NumberEval(Math.exp(mu + sd * normInv(p)));
    }

    // ---- NEGBINOM.DIST / NEGBINOMDIST / HYPGEOM.DIST / HYPGEOMDIST ----

    /** 負の二項分布: number_f 回の失敗の後に number_s 回目の成功が来る確率。 */
    static ValueEval negBinom(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 4 : a.length != 3) return ErrorEval.VALUE_INVALID;
        int f = (int) trunc(num(a[0], ec)), s = (int) trunc(num(a[1], ec));
        double p = num(a[2], ec);
        if (f < 0 || s < 1 || p <= 0 || p >= 1) return ErrorEval.NUM_ERROR; // p = 1 も #NUM! (Excel で実測)
        double pmf = chooseSmall(f + s - 1, f) * Math.pow(p, s) * Math.pow(1 - p, f);
        double cum = hasCumulative ? num(a[3], ec) : 0;
        if (!hasCumulative || cum == 0) return new NumberEval(pmf);
        // 累積 = I_p(s, f+1) (足し算より桁が落ちない)
        return new NumberEval(betI(p, s, f + 1d));
    }

    static ValueEval hyperGeom(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 5 : a.length != 4) return ErrorEval.VALUE_INVALID;
        int k = (int) trunc(num(a[0], ec)), n = (int) trunc(num(a[1], ec));
        int pop = (int) trunc(num(a[2], ec)), N = (int) trunc(num(a[3], ec));
        if (k < 0 || n < 0 || pop < 0 || N <= 0 || pop > N || n > N)
            return ErrorEval.NUM_ERROR;
        // k が取り得る範囲を超えていても Excel は #NUM! にせず、確率 0 (累積なら 1) を返す
        boolean cdf = hasCumulative && num(a[4], ec) != 0;
        int lo = cdf ? Math.max(0, n - (N - pop)) : k; // 取り得ない下限まで下げる (k が小さいほど成功が少ない)
        checkSteps((double) k - lo);
        double s = 0;
        for (int i = lo; i <= k; i++) {
            if (n - i < 0 || n - i > N - pop) continue; // あり得ない項は 0 (Excel も 0 を足す)
            s += chooseSmall(pop, i) * chooseSmall(N - pop, n - i) / chooseSmall(N, n);
        }
        return new NumberEval(s);
    }

    // ---- EXPON.DIST / EXPONDIST / WEIBULL.DIST / WEIBULL ----

    static ValueEval exponDist(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 3 : a.length != 2) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), lam = num(a[1], ec);
        double cum = hasCumulative ? num(a[2], ec) : 1;
        if (x < 0 || lam <= 0) return ErrorEval.NUM_ERROR;
        if (cum == 0) return new NumberEval(lam * Math.exp(-lam * x));
        return new NumberEval(1 - Math.exp(-lam * x));
    }

    static ValueEval weibull(ValueEval[] a, OperationEvaluationContext ec, boolean hasCumulative) throws EvaluationException {
        if (hasCumulative ? a.length != 4 : a.length != 4) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), al = num(a[1], ec), be = num(a[2], ec), cum = num(a[3], ec);
        if (x < 0 || al <= 0 || be <= 0) return ErrorEval.NUM_ERROR;
        double z = x / be;
        if (cum == 0) {
            if (x == 0) return new NumberEval(0); // Excel は x = 0 で常に 0
            double pdf = al / be * Math.pow(z, al - 1) * Math.exp(-Math.pow(z, al));
            if (!Double.isFinite(pdf)) return ErrorEval.NUM_ERROR;
            return new NumberEval(pdf);
        }
        return new NumberEval(1 - Math.exp(-Math.pow(z, al)));
    }

    // ---- T.INV / TINV / T.INV.2T ----

    static ValueEval tInvFn(ValueEval[] a, OperationEvaluationContext ec, boolean probIsLeftTail) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), df = trunc(num(a[1], ec));
        if (df < 1) return ErrorEval.NUM_ERROR;
        // 自由度 1 で確率 0 だけ #DIV/0! (Excel)。確率 1 は T.INV が #NUM!、両側は 0。両側は 1 を超えて 2 未満なら負の値になる
        if (p == 0 && df == 1) return ErrorEval.DIV_ZERO;
        if (probIsLeftTail ? (p <= 0 || p >= 1) : (p <= 0 || p >= 2)) return ErrorEval.NUM_ERROR;
        if (probIsLeftTail) return new NumberEval(tInv(p, df)); // T.INV は下側確率を渡す
        return new NumberEval(tInv2T(p, df)); // TINV・T.INV.2T は両側確率 (1 を超えると負)
    }

    // ---- Z.TEST / CONFIDENCE.NORM / CONFIDENCE.T ----

    static ValueEval zTest(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 2 || a.length > 3) return ErrorEval.VALUE_INVALID;
        List<Double> v = numbers(a[0]);
        double x = num(a[1], ec);
        int n = v.size();
        if (n == 0) return ErrorEval.NA;
        double sd;
        if (omitted(a, 2)) {
            if (n < 2) return ErrorEval.DIV_ZERO;
            sd = Math.sqrt(variance(v, n));
        } else {
            sd = num(a[2], ec);
            if (sd <= 0) return ErrorEval.NUM_ERROR;
        }
        if (sd == 0) return ErrorEval.DIV_ZERO;
        double m = 0;
        for (double d : v) m += d;
        m /= n;
        double z = (m - x) * Math.sqrt(n) / sd; // Excel は絶対値を取らない (平均 > x の片側検定)
        return new NumberEval(1 - normCdf(z));
    }

    static ValueEval confidence(ValueEval[] a, OperationEvaluationContext ec, boolean useT) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double al = num(a[0], ec), sd = num(a[1], ec);
        int n = (int) trunc(num(a[2], ec));
        if (al <= 0 || al >= 1 || sd <= 0) return ErrorEval.NUM_ERROR;
        if (n < 1) return ErrorEval.NUM_ERROR;
        if (useT && n < 2) return ErrorEval.DIV_ZERO; // CONFIDENCE.T の n=1 は #DIV/0! (Excel で実測。t 分布の自由度 0)
        double crit = useT ? tInv(1 - al / 2, n - 1) : normInv(1 - al / 2);
        return new NumberEval(sd * crit / Math.sqrt(n));
    }

    // ---- NORM.S.DIST / NORM.S.INV / NORM.DIST / NORM.INV ----
    // cumulative に TRUE/FALSE だけでなく数値 0・1 も受け付ける (Excel は 0 で密度を返す。POI 内蔵版は数値を弾く)

    static ValueEval normSDist(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double z = num(a[0], ec);
        if (num(a[1], ec) == 0) return new NumberEval(Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI)); // 確率密度
        return new NumberEval(normCdf(z));
    }

    static ValueEval normSInv(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec);
        if (p <= 0 || p >= 1) return ErrorEval.NUM_ERROR; // 0 と 1 は #NUM! (Excel)
        return new NumberEval(normInv(p));
    }

    static ValueEval normDist(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 4) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec), mu = num(a[1], ec), sd = num(a[2], ec);
        if (sd <= 0) return ErrorEval.NUM_ERROR;
        double z = (x - mu) / sd;
        if (num(a[3], ec) == 0)
            return new NumberEval(Math.exp(-z * z / 2) / (sd * Math.sqrt(2 * Math.PI))); // 確率密度
        return new NumberEval(normCdf(z));
    }

    static ValueEval normInvFn(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double p = num(a[0], ec), mu = num(a[1], ec), sd = num(a[2], ec);
        if (p <= 0 || p >= 1 || sd <= 0) return ErrorEval.NUM_ERROR; // 標準偏差 0 も #NUM! (NORM.INV(0.5,10,0) を実 Excel で確認)
        return new NumberEval(mu + sd * normInv(p));
    }


}
