package xlsx2pdf.render.formula;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 財務の関数 (減価償却・XNPV・XIRR・EFFECT・NOMINAL・PDURATION・RRI・CUMIPMT・CUMPRINC・IPMT・PPMT・ISPMT)。 */
final class FinanceFunctions {
    private FinanceFunctions() {
    }

    /**
     * ISPMT(利率, 期, 期の総数, 現在価値) = -利率 × 現在価値 × (1 - 期 / 期の総数)。
     * 元本を定額で返していくとみなした残高に利率を掛けたもの。期が範囲外でもエラーにならず値を返す。
     */
    static ValueEval ispmt(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 4) return ErrorEval.VALUE_INVALID;
        double r = num(a[0], ec), per = num(a[1], ec), nper = num(a[2], ec), pv = num(a[3], ec);
        if (nper == 0) return ErrorEval.DIV_ZERO;
        double v = -r * pv * (1 - per / nper);
        return Double.isNaN(v) || Double.isInfinite(v) ? ErrorEval.NUM_ERROR : new NumberEval(v);
    }


    // ---- 減価償却 (SLN / SYD / DDB / DB / VDB) ----

    static ValueEval sln(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double cost = num(a[0], ec), salvage = num(a[1], ec), life = num(a[2], ec);
        if (life == 0) return ErrorEval.DIV_ZERO;
        return new NumberEval((cost - salvage) / life);
    }

    static ValueEval syd(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 4) return ErrorEval.VALUE_INVALID;
        double cost = num(a[0], ec), salvage = num(a[1], ec), life = num(a[2], ec), per = num(a[3], ec);
        // 原価が 0 でも負でも値を返す。残存価額が負だと #NUM!
        if (life <= 0 || per <= 0 || salvage < 0) return ErrorEval.NUM_ERROR;
        if (per > life) return ErrorEval.NUM_ERROR; // 期が寿命を超えると #NUM!
        return new NumberEval((cost - salvage) * (life - per + 1) / (life * (life + 1) / 2));
    }

    static ValueEval ddb(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 4 && a.length != 5) return ErrorEval.VALUE_INVALID;
        double cost = num(a[0], ec), salvage = num(a[1], ec), life = num(a[2], ec), period = num(a[3], ec);
        double factor = a.length == 5 ? num(a[4], ec) : 2.0;
        if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || factor <= 0) return ErrorEval.NUM_ERROR;
        if (cost == 0) return new NumberEval(0); // DDB(0,0,5,1) は Excel でも 0
        if (period > life) return ErrorEval.NUM_ERROR; // Excel は寿命を超えた期を #NUM! にする (DDB(2400,300,10,10.5))
        if (period < 1) period = 1; // Excel は 1 年未満の期間でも第 1 期の金額をまるごと返す (DDB(1000,0,5,0.5)=400)
        // 実測 60 点から特定した式 (12 桁一致)。Excel は「期間 p の金額」を区間の累計の差ではなく
        // 時点 p-1 の簿価に利率を掛けた値として出す。期末に簿価が残存価額へ近づくほど残価が上限になる。
        //   bv   = cost * (1 - factor/life)^(period - 1)
        //   金額 = min(bv * factor/life, bv - salvage)
        //   DDB(2400,300,10,9.5) = 60.1439563122924 = 360.143956312292 - 300   (上限が効く)
        //   DDB(2400,300,10,9.3) = 75.3161530693673 = 376.580765346836 * 0.2   (利率が効く)
        //   DDB(1000,0,5,1.5)    = 309.838667696593 = 1000 * 0.6^0.5 * 0.4
        //   DDB(100000,0,5.5,5.5) = 4757.09038393337 = 100000 * (24/55)^4.5 * (4/11)
        double rate = factor / life;
        double bv = cost * Math.pow(1 - rate, period - 1);
        double one = Math.min(bv * rate, bv - salvage);
        if (one < 0)
            one = 0;
        return Double.isNaN(one) || Double.isInfinite(one) ? ErrorEval.NUM_ERROR : new NumberEval(one);
    }

    static ValueEval vdb(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 5 || a.length > 7) return ErrorEval.VALUE_INVALID;
        double cost = num(a[0], ec), salvage = num(a[1], ec), life = num(a[2], ec);
        double start = num(a[3], ec), end = num(a[4], ec);
        double factor = a.length >= 6 ? num(a[5], ec) : 2.0;
        boolean noSwitch = a.length == 7 && num(a[6], ec) != 0;
        if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < 0 || factor < 0) return ErrorEval.NUM_ERROR;
        if (cost == 0) return new NumberEval(0); // VDB(0,0,5,1,2) は Excel でも 0
        // factor = 0 は定額法へ全期切り替わるので、寿命まで損耗すれば 原価-残存価額 になる
        // (VDB(2400,300,10,0,10,FALSE) = 2100 と実 Excel で確認。FALSE は係数 0 として扱われる)
        if (start > end) return ErrorEval.NUM_ERROR; // Excel は start_period > end_period を #NUM! にする (符号反転にはしない)
        if (end > life) return ErrorEval.NUM_ERROR; // 寿命を超えた期間も #NUM! (VDB(100000,0,5,0,7))
        // 定額法への切り替えがあり、開始が小数の期のときは、Excel は累計の差でなく、開始時点の簿価から数え直す
        // (期の区切りが開始点から始まり、残りの期数も開始点から数える)
        if (!noSwitch && start != Math.floor(start)) {
            double book = cost - decliningCum(cost, salvage, life, start, factor, true, true);
            return new NumberEval(restarted(book, salvage, life, end - start, factor, life - start));
        }
        double v = decliningCum(cost, salvage, life, end, factor, !noSwitch, true)
                - decliningCum(cost, salvage, life, start, factor, !noSwitch, true);
        return new NumberEval(v);
    }

    /**
     * 簿価 book から total 期ぶんの減価額 (期の長さは 1 で、最後だけ端数)。各期は、定率法の金額と
     * 「(簿価 - 残存価額) / 残りの期数」の大きい方 (残りの期数は remaining から 1 期ごとに減る)。
     */
    static double restarted(double book, double salvage, double life, double total, double factor, double remaining) {
        checkSteps(total);
        double out = 0;
        for (int i = 1; total > 1e-12; i++) {
            double step = Math.min(1, total), avail = Math.max(book - salvage, 0);
            double per = Math.min(book * factor / life, avail);
            double rem = remaining - (i - 1);
            if (rem > 0 && avail / rem > per) per = avail / rem;
            double d = Math.min(per * step, avail);
            out += d;
            book -= step >= 1 ? per : d;
            total -= step;
        }
        return out;
    }

    /** 定額法 (一定率) の DB。記事の式: 率 = 1 − (salvage/cost)^(1/life) を小数第 3 位まで丸める。 */
    static ValueEval db(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 4 && a.length != 5) return ErrorEval.VALUE_INVALID;
        double cost = num(a[0], ec), salvage = num(a[1], ec), life = num(a[2], ec), period = num(a[3], ec);
        double month = a.length == 5 ? num(a[4], ec) : 12.0;
        int mon = (int) Math.floor(month); // Excel は月数を切り捨てる (DB(1000,0,5,1,1.5)=83.33=1000/12, 12.5 は 12)
        if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || mon < 1 || mon > 12) return ErrorEval.NUM_ERROR;
        // month が 12 のときは寿命まで、12 未満のときは寿命+1 期までが有效 (Excel で実測)
        int n = (int) period; // Excel は期間も切り捨てる (DB(1000,0,5,5.5) は 5 期分を返す)
        if (n > life + (mon < 12 ? 1 : 0)) return ErrorEval.NUM_ERROR;
        if (cost == 0) return new NumberEval(0); // 率が決められない。償却額は 0 にしておく (Excel で 0 を確認)
        double rate = Math.round((1 - Math.pow(salvage / cost, 1 / life)) * 1000) / 1000d; // 記事の「3 桁丸め」
        checkSteps(n);
        double acc = 0;
        for (int p = 1; p < n; p++) acc += fixedBalancePeriod(cost, life, rate, p, acc, mon);
        return new NumberEval(fixedBalancePeriod(cost, life, rate, n, acc, mon));
    }

    /** month は切り捨て後の整数。1 期目は月数で按分、最後の期だけ端数月を掛ける。 */
    static double fixedBalancePeriod(double cost, double life, double rate, int p, double acc, int month) {
        if (p == 1) return cost * rate * month / 12; // 最初の期だけ月数で按分
        if (month < 12 && p == (int) life + 1) return (cost - acc) * rate * (12 - month) / 12; // 最後の期だけ端数月
        return (cost - acc) * rate;
    }

    /**
     * 定率法の累計減価額。switchToStraightLine が true なら、定額法での金額が定率法より大きくなった期から定額法に切り替える (VDB の既定)。
     * 端数期間の扱いは Excel も DDB と VDB で違う: linearFraction=true ならその期の金額×割合 (VDB)、
     * false なら簿価に 1-(1-factor/life)^割合 を掛ける指数計算 (DDB: 100000*0.6^0.5-100000*0.6^1.5 = 30983.86677)。
     * どの期も「簿価が残存価額を下回らない」金額に上限される (DDB の記事の Min(...) と同じ条件)。
     */
    static double decliningCum(double cost, double salvage, double life, double t, double factor,
            boolean switchToStraightLine, boolean linearFraction) {
        if (t <= 0) return 0;
        double acc = 0;
        long whole = (long) Math.floor(t);
        checkSteps(Math.min(whole, life));
        for (long p = 1; p <= whole && p <= life; p++) {
            acc += decliningAmount(cost, salvage, life, (int) p, acc, factor, switchToStraightLine);
        }
        double frac = t - Math.floor(t);
        long next = whole + 1;
        if (frac > 0 && next <= life) {
            double book = cost - acc;
            double amt = linearFraction
                    ? decliningAmount(cost, salvage, life, (int) next, acc, factor, switchToStraightLine) * frac
                    : book * (1 - Math.pow(1 - factor / life, frac)); // Excel の DDB は指数で割る
            double cap = book - salvage; // 簿価が残存価額を下回らない
            if (amt > cap) amt = cap;
            acc += Math.max(0, amt);
        }
        return acc;
    }

    static double decliningAmount(double cost, double salvage, double life, int p, double acc,
            double factor, boolean switchToStraightLine) {
        double book = cost - acc;
        double amount = book * factor / life;
        if (switchToStraightLine) {
            double remaining = life - p + 1; // 当期を含めた残り期数
            if (remaining > 0) {
                double straightLine = (book - salvage) / remaining;
                if (straightLine > amount) amount = straightLine;
            }
        }
        double cap = book - salvage; // 簿価が残存価額を下回らない (DDB の Min と同じ)
        if (amount > cap) amount = cap;
        return amount < 0 ? 0 : amount;
    }


    // ---- XNPV / XIRR: 日付付きの正味現在価値・内部収益率 ----

    /**
     * 日付列の規則:
     *  ・XNPV は 0 以下の日付で #NUM! ({0;0}・{-5;44927}・{44927;-5})。XIRR は 0 を許す ({0;365;730} は 0.130662…) が
     *    負の日付では #NUM! ({-1;365;730})。
     *  ・どちらも減少順は #NUM! ({44927;44926})、同值是許す ({44927;44927} と {44927;44927;44927} は計算できる)。
     */
    static boolean datesOk(List<Double> d, boolean xnpv) {
        for (double x : d) {
            if (Double.isNaN(x) || Double.isInfinite(x))
                return false;
            if (xnpv ? x <= 0 : x < 0)
                return false;
        }
        // XNPV は日付が昇順でないと #NUM!。XIRR は順序の制限がない (負の日付だけ #NUM!)
        for (int i = 1; xnpv && i < d.size(); i++)
            if (d.get(i) < d.get(i - 1))
                return false;
        return true;
    }

    static ValueEval xnpv(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 3 || a.length > 4)
            return ErrorEval.VALUE_INVALID;
        double rate = num(a[0], ec);
        if (rate <= 0)
            return ErrorEval.NUM_ERROR; // 利率 0 以下は #NUM! (Excel)
        List<Double> v = numbers(a[1]), d = numbers(a[2]);
        if (v.size() != d.size() || v.isEmpty())
            return ErrorEval.NUM_ERROR; // 値と日付の個数が違うと #NUM!
        if (!datesOk(d, true))
            return ErrorEval.NUM_ERROR;
        double d0 = d.get(0), s = 0;
        for (int i = 0; i < v.size(); i++)
            s += v.get(i) / Math.pow(1 + rate, (d.get(i) - d0) / 365d);
        return Double.isNaN(s) || Double.isInfinite(s) ? ErrorEval.NUM_ERROR : new NumberEval(s);
    }

    /** XIRR の点の差の閾値 1e-8 × (|r| + 0.5)。 */
    private static double xirrTol(double r) {
        return 1e-8 * (Math.abs(r) + 0.5);
    }

    /** XIRR の f が無限大・NaN になった (Excel は #NUM!)。 */
    private static final class XirrFail extends RuntimeException {
        XirrFail() {
            super(null, null, false, false);
        }
    }

    /** 区間探しの結果。ret なら x を答えとして返す。そうでなければ [lo, hi] と、二分法の「前の中点」の初期値 prev、見つけた歩数 steps。 */
    private record XirrScan(boolean ret, double x, double lo, double hi, double prev, int steps) {
    }

    /** XIRR の区間探しの種類: 答えを返さない・正の側の 1/2 倍・負の側の 1/2 倍。 */
    private enum XirrMode { NONE, POS_HALF, NEG_HALF }

    /** XIRR の区間探しで使う値: 最初の値 v0、f(0)、推定値 g、最も早い日付の値の合計 vInf (r → +∞ での f の符号)。 */
    private record XirrCtx(java.util.function.DoubleUnaryOperator f, double v0, double f0, double g, double vInf) {
    }

    /**
     * x から step で動かして f の符号が変わる区間を探す (100 歩まで)。1 歩目以外で点の差が 1e-8 × (|点| + 0.5) より小さくなったら、
     * 正の側の 1/2 倍は f(0)/v0 が 2 以上か、0 以下で |f| < 1e-4 × |v0| (f(0) がちょうど 0 なら点の f が v0 と同じ符号) ならその点を、
     * 負の側の 1/2 倍は |f| < 2e-8 × |v0| (f(0) がちょうど 0 なら点の f が v0 と逆の符号) ならその点を答えにする。
     * 符号が変わっても区間の向きが違えば、この探しは失敗 (null)。f が無限大・NaN になったときも、この探しだけが失敗。
     */
    private static XirrScan xirrScan(XirrCtx c, double x, java.util.function.DoubleUnaryOperator step, XirrMode mode) {
        java.util.function.DoubleUnaryOperator f = c.f();
        double v0 = c.v0(), f0 = c.f0();
        try {
            for (int i = 0; i < 100; i++) {
                double nx = step.applyAsDouble(x);
                boolean conv = Math.abs(nx - x) < xirrTol(nx);
                if (conv && i > 0 && v0 != 0) {
                    if (mode == XirrMode.POS_HALF) {
                        double q = f0 / v0;
                        if (q >= 2)
                            return new XirrScan(true, nx, 0, 0, 0, 0);
                        if (q <= 0) {
                            double fn = f.applyAsDouble(nx);
                            if (Math.abs(fn) < 1e-4 * Math.abs(v0) && (q < 0 || fn / v0 > 0))
                                return new XirrScan(true, nx, 0, 0, 0, 0);
                        }
                    }
                    if (mode == XirrMode.NEG_HALF) {
                        double fn = f.applyAsDouble(nx);
                        if (Math.abs(fn) < 2e-8 * Math.abs(v0) && (f0 != 0 || fn / v0 < 0))
                            return new XirrScan(true, nx, 0, 0, 0, 0);
                    }
                }
                if ((f.applyAsDouble(nx) > 0) != (f.applyAsDouble(x) > 0)) {
                    double lo = Math.min(x, nx), hi = Math.max(x, nx);
                    return xirrWrongOrientation(c, lo, hi) ? null : new XirrScan(false, 0, lo, hi, nx, i + 1);
                }
                x = nx;
            }
            return null;
        } catch (XirrFail e) {
            return null;
        }
    }

    /**
     * 符号の変わる区間 [lo, hi] の向きが違うか (Excel はその区間を使わない)。-1 より下の区間ではどちらの判定も逆にする。
     * <ul>
     * <li>下端の f が v0 と同じ符号 (v0 が負なら f が増えていく向き)。</li>
     * <li>正の側、または推定値が正のとき: 0 に近い端の f が f(0) と逆の符号 (区間と 0 のあいだにも符号の変化がある)。</li>
     * <li>-1 より下の区間で、最も早い日付の値の合計 (r → +∞ での f の符号) が v0 と逆の符号。</li>
     * </ul>
     */
    private static boolean xirrWrongOrientation(XirrCtx c, double lo, double hi) {
        java.util.function.DoubleUnaryOperator f = c.f();
        boolean below = hi < -1;
        double near = hi <= 0 ? hi : lo;
        boolean wrong = false;
        if (lo >= 0 || c.g() > 0)
            wrong = ((f.applyAsDouble(near) > 0) != (c.f0() > 0)) != below;
        if (c.v0() != 0) {
            wrong |= ((f.applyAsDouble(lo) > 0) != (c.v0() < 0)) != below;
            if (below && (c.vInf() > 0) != (c.v0() > 0))
                wrong = true;
        }
        return wrong;
    }

    /** b の n 乗 (n ≥ 0) を、下位のビットから 2 乗を繰り返して求める (Excel の XIRR の整数年の累乗)。 */
    private static double xirrIntPow(double b, long n) {
        double r = 1, x = b;
        while (n != 0) {
            if ((n & 1) != 0)
                r *= x;
            x *= x;
            n >>= 1;
        }
        return r;
    }

    /**
     * XIRR(値, 日付, [推定値])。Excel 2024 に約 1 万 4 千式を計算させて決めた規則 (日付を並べ替えて利回りが -100% より下になる
     * 一部の式を除いて、推定値 0・負の利回り・符号が何度も変わる流れ・#NUM! を含めて最後の桁まで一致):
     * <ul>
     * <li>値に正と負の両方がなければ #NUM!。推定値が -1 以下も #NUM!。推定値 0 は 1e-5 として扱う。日付は整数に切り捨てる。</li>
     * <li>f(r) = Σ v_i / (1+r)^((d_i − d_0)/365)。累乗は、日数が 365 の倍数なら 2 乗の繰り返し (負なら逆数)、それ以外は
     * exp(日数/365 × log(1+r)) (1+r ≤ 0 なら NaN)。2 番目以降の項を前から足し、最初の値を最後に足す。-1 で打ち切らない。</li>
     * <li>推定値が正: 2 倍ずつ、次に 1/2 倍ずつ動かして符号の変わる区間を探す (xirrScan)。見つからなければ -|推定値| から 1/2 倍、
     * 次に -1 との中点 (x − (1 + x) / 2) へ進めて探す (-1 より下からは中点が先)。</li>
     * <li>区間を hi − (hi − lo) / 2 で半分にしていき、前の中点 (最初は区間探しの最後の点) との差が 1e-8 × (|中点| + 0.5) より小さく、
     * 中点が負なら |f(中点)| < 2e-8 × |v0| も満たした中点を返す。区間を見つけた歩数と二分法の回数の合計が 98 を超えたら #NUM!。
     * 二分法の途中で f が無限大・NaN になっても #NUM!。</li>
     * </ul>
     */
    static ValueEval xirr(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 2 || a.length > 3)
            return ErrorEval.VALUE_INVALID;
        List<Double> v = numbers(a[0]), d = numbers(a[1]);
        if (v.size() != d.size() || v.isEmpty())
            return ErrorEval.NUM_ERROR; // 値と日付の個数が違うと #NUM!
        if (!datesOk(d, false))
            return ErrorEval.NUM_ERROR;
        double guess = a.length == 3 ? num(a[2], ec) : 0.1;
        if (Double.isNaN(guess) || guess <= -1)
            return ErrorEval.NUM_ERROR;
        boolean anyPos = false, anyNeg = false;
        for (double x : v) {
            anyPos |= x > 0;
            anyNeg |= x < 0;
        }
        if (!anyPos || !anyNeg)
            return ErrorEval.NUM_ERROR;
        if (guess == 0)
            guess = 1e-5;
        int n = v.size();
        double[] vals = new double[n], days = new double[n];
        double d0 = Math.floor(d.get(0)), dMin = Double.POSITIVE_INFINITY;
        for (int i = 0; i < n; i++) {
            vals[i] = v.get(i);
            days[i] = Math.floor(d.get(i)) - d0;
            dMin = Math.min(dMin, days[i]);
        }
        double vInf = 0;
        for (int i = 0; i < n; i++)
            if (days[i] == dMin)
                vInf += vals[i];
        java.util.function.DoubleUnaryOperator f = r -> {
            double b = 1 + r, s = 0;
            for (int k = 1; k <= n; k++) {
                int i = k % n; // 2 番目以降を前から、最初の値を最後に
                double dy = days[i], p;
                if (dy % 365 == 0) {
                    long m = (long) (dy / 365);
                    p = m >= 0 ? xirrIntPow(b, m) : 1 / xirrIntPow(b, -m);
                } else {
                    p = b > 0 ? Math.exp(dy / 365 * Math.log(b)) : Double.NaN;
                }
                s += vals[i] / p;
            }
            if (!Double.isFinite(s))
                throw new XirrFail();
            return s;
        };
        double v0 = vals[0];
        try {
            XirrCtx c = new XirrCtx(f, v0, f.applyAsDouble(0), guess, vInf);
            XirrScan sc = null;
            if (guess > 0) {
                sc = xirrScan(c, guess, x -> x * 2, XirrMode.NONE);
                if (sc == null)
                    sc = xirrScan(c, guess, x -> x * 0.5, XirrMode.POS_HALF);
                if (sc != null && sc.ret())
                    return new NumberEval(sc.x());
            }
            if (sc == null) {
                double x = -Math.abs(guess);
                java.util.function.DoubleUnaryOperator half = y -> y * 0.5, mid = y -> y - (1 + y) / 2;
                if (x < -1) {
                    sc = xirrScan(c, x, mid, XirrMode.NONE);
                    if (sc == null)
                        sc = xirrScan(c, x, half, XirrMode.NEG_HALF);
                } else {
                    sc = xirrScan(c, x, half, XirrMode.NEG_HALF);
                    if (sc == null)
                        sc = xirrScan(c, x, mid, XirrMode.NONE);
                }
            }
            if (sc == null)
                return ErrorEval.NUM_ERROR;
            if (sc.ret())
                return new NumberEval(sc.x());
            double lo = sc.lo(), hi = sc.hi(), prev = sc.prev();
            boolean loPositive = f.applyAsDouble(lo) > 0;
            for (int k = sc.steps() + 1; k <= 98; k++) {
                double m = hi - (hi - lo) / 2;
                double fm = f.applyAsDouble(m);
                if (Math.abs(m - prev) < xirrTol(m) && (m >= 0 || Math.abs(fm) < 2e-8 * Math.abs(v0)))
                    return new NumberEval(m);
                prev = m;
                if ((fm > 0) == loPositive)
                    lo = m;
                else
                    hi = m;
            }
            return ErrorEval.NUM_ERROR;
        } catch (XirrFail e) {
            return ErrorEval.NUM_ERROR;
        }
    }

    // ---- EFFECT / NOMINAL: 名目金利と実効金利 ----

    static ValueEval effect(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2)
            return ErrorEval.VALUE_INVALID;
        double nom = num(a[0], ec), per = num(a[1], ec);
        int n = (int) Math.floor(per); // Excel は複利回数を切り捨てる
        if (nom <= 0 || n < 1)
            return ErrorEval.NUM_ERROR;
        return new NumberEval(Math.pow(1 + nom / n, n) - 1);
    }

    static ValueEval nominal(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 2)
            return ErrorEval.VALUE_INVALID;
        double eff = num(a[0], ec), per = num(a[1], ec);
        int n = (int) Math.floor(per);
        if (eff <= 0 || n < 1)
            return ErrorEval.NUM_ERROR;
        return new NumberEval(n * (Math.pow(1 + eff, 1d / n) - 1));
    }

    // ---- PDURATION / RRI: 期間と成長率 ----

    static ValueEval pduration(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3)
            return ErrorEval.VALUE_INVALID;
        double rate = num(a[0], ec), pv = num(a[1], ec), fv = num(a[2], ec);
        if (rate <= 0 || pv <= 0 || fv <= 0)
            return ErrorEval.NUM_ERROR;
        double v = Math.log(fv / pv) / Math.log(1 + rate);
        return Double.isNaN(v) || Double.isInfinite(v) ? ErrorEval.NUM_ERROR : new NumberEval(v);
    }

    static ValueEval rri(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3)
            return ErrorEval.VALUE_INVALID;
        double nper = num(a[0], ec), pv = num(a[1], ec), fv = num(a[2], ec);
        // 実測: RRI(10,100,0) = -1 (fv=0 を許す)、RRI(10,-100,-200) = 0.0717734625362931 (pv・fv がともに負でも比が正なら可)、
        //       RRI(10,-100,200)・RRI(10,0,200)・RRI(0,100,100) は #NUM!。
        if (nper <= 0 || pv == 0 || fv / pv < 0)
            return ErrorEval.NUM_ERROR;
        double v = Math.pow(fv / pv, 1 / nper) - 1;
        return Double.isNaN(v) || Double.isInfinite(v) ? ErrorEval.NUM_ERROR : new NumberEval(v);
    }

    // ---- CUMIPMT / CUMPRINC: 区間の利払い・元本返済の累計 ----

    static ValueEval cumipmt(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        return cumipmtLike(a, ec, false);
    }

    static ValueEval cumprinc(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        return cumipmtLike(a, ec, true);
    }

    /**
     * 償還の規則:
     *  ・type=0 (期末): 利子 = 期首日順残高 × rate。期首残高から元本を減らす。
     *  ・type=1 (期首): 1 期目の利子は 0。以降は同じ式 (実測 CUMIPMT(2,3,1) = -164.135389 に対応)。
     *  ・nper は小数のまま払込額に出す (実測 CUMIPMT(0.1,10.9,...) に対応。切捨てしない)。
     *  ・利子の総和は「払込額 × 期数 − 残高の増分」で閉じた形になる (Σ元本 = 残高の減少分)。
     * Excel は pv>0 のとき利子・元本ともに負で返す (CUMIPMT+CUMPRINC = 期数 × PMT)。
     */
    static double amortPay(double rate, double nper, double pv, double fv, int type) {
        double pay = rate == 0 ? (pv + fv) / nper
                : (pv * rate + fv / Math.pow(1 + rate, nper)) / (1 - Math.pow(1 + rate, -nper));
        return type == 1 ? pay / (1 + rate) : pay;
    }

    /** k 期の払込みが終わった時点の残高 (pv - 元本返済累計)。 */
    static double balance(double rate, double nper, double pv, double fv, int type, int k, double pay) {
        if (k <= 0)
            return pv;
        if (rate == 0)
            return pv - k * pay;
        if (type == 0) {
            double q = Math.pow(1 + rate, k);
            return (pv - pay / rate) * q + pay / rate;
        }
        // type=1: B_1 = pv - pay (1 期目に利子が付かない)、B_k = B_{k-1}(1+rate) - pay
        double p = Math.pow(1 + rate, k - 1);
        return (pv - pay) * p - pay * (p - 1) / rate;
    }

    static ValueEval cumipmtLike(ValueEval[] a, OperationEvaluationContext ec, boolean principal)
            throws EvaluationException {
        if (a.length != 6)
            return ErrorEval.VALUE_INVALID;
        double rate = num(a[0], ec), nd = num(a[1], ec), pv = num(a[2], ec);
        double sd = num(a[3], ec), ed = num(a[4], ec), td = num(a[5], ec);
        double n = Math.floor(nd);
        // 生の値での範囲チェック (実測: start=stop=10.5 と stop=10.5 は #NUM! なのに、9.5 と 2.5 は効く
        // → 切り捨てた値では見ていない)。start > stop も生の値で見る (実測 (5,3) = #NUM!)。
        if (rate <= 0 || n < 1 || pv <= 0 || sd < 1 || ed < 1 || sd > nd || ed > nd || sd > ed
                || (td != 0 && td != 1))
            return ErrorEval.NUM_ERROR;
        // 実測: start は切り上げ、stop は切り捨て。区間が空 (ceil(start) > floor(stop)) なら 0
        //  (1.2,1.8)=0 (1.1,1.1)=0 (1.5,1.5)=0 (2.5,2.5)=0、(1.5,2)=-93.725=P2、(9.5,10)=-14.795=P10)。
        int start = (int) Math.ceil(sd), end = (int) Math.floor(ed);
        if (end < start)
            return new NumberEval(0);
        double pay = amortPay(rate, nd, pv, 0, (int) td);
        double b0 = balance(rate, nd, pv, 0, (int) td, start - 1, pay);
        double b1 = balance(rate, nd, pv, 0, (int) td, end, pay);
        double princ = b0 - b1; // 区間の元本返済
        double v = principal ? -princ : -((end - start + 1) * pay - princ); // 利子 = 支払 - 元本
        return Double.isNaN(v) || Double.isInfinite(v) ? ErrorEval.NUM_ERROR : new NumberEval(v);
    }

    // ---- IPMT / PPMT / DELTA / ROMAN: POI 版が 1 期目で #VALUE! になる・既定引数が効かないので自前で実装 ----

    static ValueEval ipmt(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        return ipmtLike(a, ec, true);
    }

    static ValueEval ppmt(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        return ipmtLike(a, ec, false);
    }

    /** IPMT・PPMT(利率, 期, 期の総数, 現在価値, [将来価値], [支払期日]) (interest=true なら利息、false なら元本)。 */
    static ValueEval ipmtLike(ValueEval[] a, OperationEvaluationContext ec, boolean interest)
            throws EvaluationException {
        if (a.length < 4 || a.length > 6)
            return ErrorEval.VALUE_INVALID;
        double rate = num(a[0], ec), per = num(a[1], ec), nper = num(a[2], ec), pv = num(a[3], ec);
        double fv = omitted(a, 4) ? 0 : num(a[4], ec);
        int type = !omitted(a, 5) && num(a[5], ec) != 0 ? 1 : 0;
        if (per < 1 || Math.floor(per) > nper || nper <= 0)
            return ErrorEval.NUM_ERROR; // 期は 1 以上で、小数は切り捨てて総数以内
        double pay = pmt(rate, nper, pv, fv, type);
        // 利息 = 利率 × 期の初めの残高 (期首払いは 1 期目が 0 で、残高を (1 + 利率) で割る)
        double ip;
        if (rate == 0)
            ip = 0;
        else if (type == 1)
            ip = per <= 1 ? 0 : -rate * balanceAfter(rate, nper, pv, fv, per - 1) / (1 + rate);
        else
            ip = -rate * balanceAfter(rate, nper, pv, fv, per - 1);
        double v = interest ? ip : pay - ip;
        return Double.isNaN(v) || Double.isInfinite(v) ? ErrorEval.NUM_ERROR : new NumberEval(v);
    }

    /**
     * k 期の支払のあとの残高 = pv q^k + pay (q^k - 1)/r を pay を代入して整理した形。引き算で桁が落ちない
     * (q = 1 + 利率、利率が大きく期が長いときも)。
     */
    static double balanceAfter(double rate, double nper, double pv, double fv, double k) {
        double q = 1 + rate, qn = Math.pow(q, nper), qk = Math.pow(q, k);
        return (pv * (qn - qk) - fv * (qk - 1)) / (qn - 1);
    }

    /** PMT (Excel の符号: 現在価値が正なら支払は負)。 */
    static double pmt(double rate, double nper, double pv, double fv, int type) {
        if (rate == 0)
            return -(pv + fv) / nper;
        double q = Math.pow(1 + rate, nper);
        return -(pv * q + fv) * rate / ((q - 1) * (1 + rate * type));
    }

    /** DELTA(数1, [数2=0]) — 第 2 引数の既定値 0 と空セルの 0 化 (実測 DELTA(1)=0, DELTA(空白)=1)。 */

}
