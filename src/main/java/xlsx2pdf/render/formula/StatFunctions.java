package xlsx2pdf.render.formula;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 統計の関数 (順位・百分位・四分位・最頻値・RSQ・STEYX・KURT・SKEW・HARMEAN・TRIMMEAN・PROB・FISHER)。 */
final class StatFunctions {
    private StatFunctions() {
    }

    // ---- PERCENTILE.INC / PERCENTILE.EXC / QUARTILE.INC / QUARTILE.EXC ----

    /** 四分位・百分位。inc は n-1 基準 (Excel の INC)、exc は n+1 基準。quart のとき引数は 1〜4。 */
    static ValueEval percentile(ValueEval[] a, OperationEvaluationContext ec, boolean inc, boolean quart) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double k;
        if (quart) {
            double q = num(a[1], ec);
            if (inc ? q < 0 || q > 4 : q < 1 || q > 3) return ErrorEval.NUM_ERROR; // INC は 0〜4 (0=最小, 4=最大)、EXC は 1〜3 (Excel)
            k = (long) q / 4.0; // quart は整数に切り捨ててから 4 で割る (Excel)
        } else {
            k = num(a[1], ec);
        }
        if (inc ? k < 0 || k > 1 : k <= 0 || k >= 1) return ErrorEval.NUM_ERROR;
        List<Double> v = numbers(a[0]);
        int n = v.size();
        if (n == 0) return ErrorEval.NUM_ERROR;
        v.sort(null);
        // 0 基準の位置にする: INC は 1 + k(n-1) - 1 = k(n-1)、EXC は k(n+1) - 1
        double p = inc ? k * (n - 1) : k * (n + 1) - 1;
        if (p < 0 || p > n - 1) return ErrorEval.NUM_ERROR;
        int i = (int) Math.floor(p);
        if (i >= n) i = n - 1;
        double f = p - i, x = v.get(i);
        if (i + 1 < n) x += f * (v.get(i + 1) - x); // 足し算の順序を Excel と同じ (1-k)・x + k・次の値 にしない
        return new NumberEval(x);
    }

    // ---- MODE.SNGL ----

    /** 最頻値。同じ個数のものが複数あるときは一番小さい方 (Excel)。1 個しか重複がなければ #N/A。 */
    static ValueEval modeSngl(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        List<Double> v = new ArrayList<>();
        for (ValueEval x : a) {
            // 直接書いた数値にならない文字列は無視する (Excel)
            if (x instanceof StringEval s && numberOrNull(s.getStringValue()) == null) continue;
            collect(x, false, v);
        }
        if (v.isEmpty()) return ErrorEval.NA;
        v.sort(null);
        int best = 0, at = 0;
        for (int i = 0; i < v.size();) {
            int j = i;
            double vi = v.get(i); // Double 同士を == で比べない (箱入れたものは数値で比べる)
            while (j < v.size() && v.get(j) == vi) j++;
            if (j - i > best) {
                best = j - i;
                at = i;
            }
            i = j;
        }
        if (best < 2) return ErrorEval.NA;
        return new NumberEval(v.get(at));
    }

    // ---- RANK.EQ / RANK.AVG ----

    /** 順位。order が 0 でない (省略除く) なら昇順。同じ値の個数を avg のとき順位の平均にする。 */
    static ValueEval rank(ValueEval[] a, OperationEvaluationContext ec, boolean avg) throws EvaluationException {
        if (a.length < 2 || a.length > 3) return ErrorEval.VALUE_INVALID;
        double x = num(a[0], ec);
        boolean asc = !omitted(a, 2) && num(a[2], ec) != 0;
        int better = 0, same = 0;
        for (double d : numbers(a[1])) {
            if (d == x) same++;
            else if (asc ? d < x : d > x) better++;
        }
        if (same == 0) return ErrorEval.NA; // 範囲に無い数値は #N/A (Excel)
        return new NumberEval(better + 1 + (avg ? (same - 1) / 2.0 : 0));
    }


    // ---- RSQ / STEYX / KURT / SKEW / HARMEAN / TRIMMEAN / PROB / FISHER ----

    /** known_y・known_x から {n, 平均 x, 平均 y, Sxx, Syy, Sxy} を作る。範囲の大きさが違うと null (#N/A)。 */
    static double[] regression(ValueEval[] a) throws EvaluationException {
        List<Double> x = positional(a[1]), y = positional(a[0]);
        if (x.size() != y.size()) return null; // 範囲の個数が違う → #N/A (Excel)
        int n = 0;
        double[] px = new double[x.size()], py = new double[y.size()];
        for (int i = 0; i < x.size(); i++) {
            double xv = x.get(i), yv = y.get(i);
            if (Double.isNaN(xv) || Double.isNaN(yv)) continue; // どちらかが数値でない組は捨てる (Excel は計算から除く)
            px[n] = xv;
            py[n] = yv;
            n++;
        }
        if (n == 0) return null;
        double mx = 0, my = 0;
        for (int i = 0; i < n; i++) {
            mx += px[i];
            my += py[i];
        }
        mx /= n;
        my /= n;
        double sxx = 0, syy = 0, sxy = 0;
        for (int i = 0; i < n; i++) {
            double dx = px[i] - mx, dy = py[i] - my;
            sxx += dx * dx;
            syy += dy * dy;
            sxy += dx * dy;
        }
        return new double[] { n, mx, my, sxx, syy, sxy };
    }

    /** 範囲を行优先で平らにする。数値でないマスは NaN なので、組単位で捨てられる。 */
    static List<Double> positional(ValueEval v) throws EvaluationException {
        List<Double> out = new ArrayList<>();
        if (v instanceof RefEval re) v = re.getInnerValueEval(re.getFirstSheetIndex());
        if (v instanceof AreaEval ar) {
            for (int r = 0; r < ar.getHeight(); r++)
                for (int c = 0; c < ar.getWidth(); c++) {
                    ValueEval x = ar.getRelativeValue(r, c);
                    out.add(x instanceof NumberEval n ? n.getNumberValue() : Double.NaN);
                }
            return out;
        }
        out.add(v instanceof NumberEval n ? n.getNumberValue() : Double.NaN);
        return out;
    }

    static ValueEval rsq(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double[] r = regression(a);
        if (r == null) return ErrorEval.NA; // known_y と known_x の個数が違う (Excel)
        if (r[3] == 0 || r[4] == 0) return ErrorEval.DIV_ZERO;
        return new NumberEval(r[5] * r[5] / (r[3] * r[4]));
    }

    static ValueEval steyx(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double[] r = regression(a);
        if (r == null) return ErrorEval.NA;
        int n = (int) r[0];
        if (n < 3 || r[3] == 0) return ErrorEval.DIV_ZERO;
        double sse = r[4] - r[5] * r[5] / r[3];
        if (sse < 0) sse = 0; // 四捨五入で少し負になったとき
        return new NumberEval(Math.sqrt(sse / (n - 2)));
    }

    static ValueEval kurt(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        List<Double> v = new ArrayList<>();
        for (ValueEval x : a) collect(x, false, v);
        int n = v.size();
        if (n < 4) return ErrorEval.DIV_ZERO;
        double m = 0;
        for (double d : v) m += d;
        m /= n;
        double s2 = 0, s4 = 0;
        for (double d : v) {
            double t = d - m;
            s2 += t * t;
            s4 += t * t * t * t;
        }
        if (s2 == 0) return ErrorEval.DIV_ZERO;
        double sd = Math.sqrt(s2 / (n - 1));
        double sum4 = s4 / (sd * sd * sd * sd);
        return new NumberEval(n * (n + 1d) / ((n - 1d) * (n - 2d) * (n - 3d)) * sum4 - 3 * (n - 1d) * (n - 1d) / ((n - 2d) * (n - 3d)));
    }

    static ValueEval skew(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        List<Double> v = new ArrayList<>();
        for (ValueEval x : a) collect(x, false, v);
        int n = v.size();
        if (n < 3) return ErrorEval.DIV_ZERO;
        double m = 0;
        for (double d : v) m += d;
        m /= n;
        double s2 = 0, s3 = 0;
        for (double d : v) {
            double t = d - m;
            s2 += t * t;
            s3 += t * t * t;
        }
        if (s2 == 0) return ErrorEval.DIV_ZERO;
        double sd = Math.sqrt(s2 / (n - 1));
        return new NumberEval(n / ((n - 1d) * (n - 2d)) * s3 / (sd * sd * sd));
    }

    static ValueEval harmean(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        List<Double> v = new ArrayList<>();
        for (ValueEval x : a) collect(x, false, v);
        if (v.isEmpty()) return ErrorEval.NA;
        double s = 0;
        for (double d : v) {
            if (d <= 0) return ErrorEval.NUM_ERROR;
            s += 1 / d;
        }
        if (s == 0) return ErrorEval.DIV_ZERO;
        return new NumberEval(v.size() / s);
    }

    static ValueEval trimMean(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2) return ErrorEval.VALUE_INVALID;
        double p = num(a[1], ec);
        List<Double> v = numbers(a[0]);
        int n = v.size();
        if (p < 0 || p >= 1) return ErrorEval.NUM_ERROR;
        if (n == 0) return ErrorEval.NUM_ERROR;
        v.sort(null);
        int cut = (int) Math.floor(n * p / 2);
        if (cut * 2 >= n) return ErrorEval.NUM_ERROR;
        double s = 0;
        for (int i = cut; i < n - cut; i++) s += v.get(i);
        return new NumberEval(s / (n - 2 * cut));
    }

    static ValueEval prob(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 3 || a.length > 5) return ErrorEval.VALUE_INVALID;
        List<Double> xs = numbers(a[0]), ps = numbers(a[1]);
        if (xs.size() != ps.size() || xs.isEmpty()) return ErrorEval.NA;
        double low = num(a[2], ec);
        double up = omitted(a, 3) ? low : num(a[3], ec);
        boolean cum = !omitted(a, 4) && num(a[4], ec) != 0;
        double tot = 0;
        for (double p : ps) {
            if (p < 0 || p > 1) return ErrorEval.NUM_ERROR;
            tot += p;
        }
        // 確率範囲の合計が 1 でないと Excel は #NUM! を返す (PROB(A306:C306,A307:C307,0.2) で実測)
        if (Math.abs(tot - 1) > 1e-8) return ErrorEval.NUM_ERROR;
        double s = 0;
        for (int i = 0; i < xs.size(); i++) {
            boolean in = cum ? xs.get(i) <= up : xs.get(i) >= low && xs.get(i) <= up;
            if (in) s += ps.get(i);
        }
        return new NumberEval(s);
    }

    static ValueEval fisher(ValueEval[] a, OperationEvaluationContext ec, boolean inverse) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        double z = num(a[0], ec);
        if (inverse) return new NumberEval((Math.exp(2 * z) - 1) / (Math.exp(2 * z) + 1));
        if (z <= -1 || z >= 1) return ErrorEval.NUM_ERROR;
        return new NumberEval(0.5 * Math.log((1 + z) / (1 - z))); // Excel と同じ式 (log1p だと 1 に近い値で 8 桁目が違う)
    }


}
