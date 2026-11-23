package xlsx2pdf.render.formula;

import static xlsx2pdf.render.formula.MathFunctions.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 分布関数の数値コア (不完全ガンマ・不完全ベータ・正規分布・t・χ2・F・二項分布の確率)。 */
final class SpecialMath {
    private SpecialMath() {
    }

    // ==================== 確率分布の数値コア ====================
    //
    // 精度の確かめ方: 正規分布は POI の NORMSDIST・NORMSINV と、t 分布は POI の TDIST と、χ2 は偶数自由度の多項式と、
    // 不完全ベータは I_x(1/2,1/2) = (2/π)asin(√x)・I_x(2,3) = 6x²-8x³+3x⁴ と比べた。
    // 逆関数は二分法 (上側確率が小さいときは上側確率を直接解く) なので、分布関数が合えば同じ桁まで合う。

    /** 正則不完全ガンマ P(a,x)。x < a+1 は級数、それ以外は連分数。 */
    static double gamP(double a, double x) {
        if (x <= 0) return 0;
        if (x < a + 1.0) return gamSeries(a, x);
        return 1 - gamCf(a, x);
    }

    /** 正則不完全ガンマ Q(a,x) = 1 - P(a,x)。 */
    static double gamQ(double a, double x) {
        if (x <= 0) return 1;
        if (x < a + 1.0) return 1 - gamSeries(a, x);
        return gamCf(a, x);
    }

    static double gamSeries(double a, double x) {
        double ap = a, sum = 1 / a, del = sum;
        for (int n = 0; n < 600; n++) {
            ap += 1;
            del *= x / ap;
            sum += del;
            if (Math.abs(del) < Math.abs(sum) * 1e-17) break;
        }
        double lg = -x + a * Math.log(x) - logGamma(a);
        return lg < -745 ? 0 : Math.exp(lg) * sum;
    }

    static double gamCf(double a, double x) {
        double b = x + 1 - a, c = 1e300, d = 1 / b, h = d;
        for (int i = 1; i <= 500; i++) {
            double an = -i * (i - a);
            b += 2;
            d = an * d + b;
            if (Math.abs(d) < 1e-300) d = 1e-300;
            c = b + an / c;
            if (Math.abs(c) < 1e-300) c = 1e-300;
            d = 1 / d;
            double del = d * c;
            h *= del;
            if (Math.abs(del - 1) < 1e-17) break;
        }
        double lg = -x + a * Math.log(x) - logGamma(a);
        return lg < -745 ? 0 : Math.exp(lg) * h;
    }

    /** 正則不完全ベータ I_x(a,b) (連分数)。x が大きい側は対称性で回る。 */
    static double betI(double x, double a, double b) {
        if (x <= 0) return 0;
        if (x >= 1) return 1;
        if (x > (a + 1) / (a + b + 2)) return 1 - betI(1 - x, b, a);
        double bt = Math.exp(a * Math.log(x) + b * Math.log1p(-x) - logBetaAB(a, b));
        double qab = a + b, qap = a + 1, qam = a - 1;
        double c = 1, d = 1 - qab * x / qap;
        if (Math.abs(d) < 1e-300) d = 1e-300;
        d = 1 / d;
        double h = d;
        for (int m = 1; m <= 500; m++) {
            int m2 = 2 * m;
            double aa = m * (b - m) * x / ((qam + m2) * (a + m2));
            d = 1 + aa * d;
            if (Math.abs(d) < 1e-300) d = 1e-300;
            c = 1 + aa / c;
            if (Math.abs(c) < 1e-300) c = 1e-300;
            d = 1 / d;
            h *= d * c;
            aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
            d = 1 + aa * d;
            if (Math.abs(d) < 1e-300) d = 1e-300;
            c = 1 + aa / c;
            if (Math.abs(c) < 1e-300) c = 1e-300;
            d = 1 / d;
            double del = d * c;
            h *= del;
            if (Math.abs(del - 1) < 1e-17) break;
        }
        return bt * h / a;
    }

    static double logBetaAB(double a, double b) {
        return logGamma(a) + logGamma(b) - logGamma(a + b);
    }

    /** 標準正規の CDF (自前の erfc で作る)。 */
    static double normCdf(double x) {
        return x < 0 ? 0.5 * erfc(-x * 0.7071067811865475) : 1 - 0.5 * erfc(x * 0.7071067811865475);
    }

    /** 標準正規の逆 CDF (中央部は二分法、極端な尾は対数-domain の漸化近似)。 */
    static double normInv(double p) {
        if (p <= 0 || p >= 1) return Double.NaN;
        if (p == 0.5) return 0;
        // 上側尾は下側尾に載せ替える。1 - p は 0.5 <= p < 1 で厳密に計算できる (Sterbenz の補題) ので、
        // 中央部の二分法で 1 - Φ(x) の打消ちを踏むより桁が残る。実測では
        //   NORM.S.INV(1E-5)〜NORM.S.INV(1E-17) の 13 点がこの下側経路ですべて Excel と 1e-14 相対で一致し、
        //   上側を直接解いていた旧経路は 1-p が 1e-8 より小さくなると劣化した (1-1E-15 で Excel 7.94144448741598 / 旧 7.9347)。
        if (p > 0.5)
            return -normInv(1 - p);
        // 1e-25 より外側は CDF の値が 0 に近付いて二分法では解けない (実測で -Infinity に発散した)。
        // Excel は NORM.S.INV(1E-300) = -37.0470962993612 を返すので、対数-domain で解く。
        if (p < 1e-25) {
            double target = Math.log(p);
            double a = 5, b = 45; // ln Q(5) ≒ -14.5, ln Q(45) ≒ -1015 なので 5e-324 まで入る
            for (int i = 0; i < 200; i++) {
                double m = (a + b) / 2;
                if (m == a || m == b) break;
                // logQ は x の増加について減少する。logQ(m) が target より大きい間は
                // 根は m より右側にある (＝下側境界 a を上げる)。
                if (logQ(m) > target)
                    a = m;
                else
                    b = m;
            }
            return -(a + b) / 2;
        }
        return solve(SpecialMath::normCdf, p, -6, 0);
    }

    /** x > 0 のときの ln Q(x) (Q = 1 - Φ)。Q(x) = φ(x)/x・(1 - 1/x² + 3/x⁴ - 15/x⁶ + …) の
     *  括弧内を対数-domain で扱う (ミルの級数は発散するので最小項で打ち切る)。 */
    static double logQ(double x) {
        double t = 1 / (x * x), s = 0, prev = 0, prod = 1;
        for (int k = 1; k <= 60; k++) {
            prod *= (2 * k - 1); // (2k-1)!!
            double mag = prod * Math.pow(t, k);
            double add = (k & 1) == 1 ? -mag : mag; // 1 - 1/x² + 3/x⁴ - … なので第 1 項は負
            if (k > 1 && Math.abs(add) > Math.abs(prev)) break; // 項が増え始める前に打ち切る (漸近級数)
            s += add;
            prev = add;
        }
        if (1 + s <= 0) s = -0.9999999999; // 打ち切りが極端に悪い場合 (x が小さいときだけ。通常は来ない)
        return -x * x / 2 - 0.9189385332046727 - Math.log(x) + Math.log1p(s); // -x²/2 - ln(x√(2π)) + ln(1 + 級数)
    }

    /** 単調な f(x) = target を lo・hi を広げながら二分法で探す (逆関数用)。 */
    static double solve(java.util.function.DoubleUnaryOperator f, double target, double lo, double hi) {
        for (int g = 0; g < 200 && f.applyAsDouble(hi) < target; g++) hi = hi * 2 + 2;
        for (int g = 0; g < 200 && f.applyAsDouble(lo) > target; g++) lo = lo * 2 - 2;
        for (int i = 0; i < 400; i++) {
            double m = (lo + hi) / 2;
            if (m == lo || m == hi) break;
            if (f.applyAsDouble(m) < target) lo = m;
            else hi = m;
        }
        return (lo + hi) / 2;
    }

    /** 自由度 df の t 分布の CDF。 */
    /** 単調に減る f(x) = target を、lo・hi を広げながら二分法で探す (上側確率の逆関数用)。 */
    static double solveDecreasing(java.util.function.DoubleUnaryOperator f, double target, double lo, double hi) {
        for (int g = 0; g < 200 && f.applyAsDouble(hi) > target; g++) hi = hi * 2 + 2;
        for (int g = 0; g < 200 && f.applyAsDouble(lo) < target; g++) lo = lo * 2 - 2;
        for (int i = 0; i < 400; i++) {
            double m = (lo + hi) / 2;
            if (m == lo || m == hi) break;
            if (f.applyAsDouble(m) > target) lo = m;
            else hi = m;
        }
        return (lo + hi) / 2;
    }

    /**
     * 下側確率 p の点。p が 0.5 を超えるときは、上側確率 1 - p を直接解く (1 に近い p を 1 から引くと、
     * 尾の確率の桁が落ちるため)。cdf は下側確率、sf は上側確率 (1 - cdf を桁を落とさずに求めたもの)。
     */
    static double invert(java.util.function.DoubleUnaryOperator cdf, java.util.function.DoubleUnaryOperator sf, double p, double lo, double hi) {
        return p > 0.5 ? solveDecreasing(sf, 1 - p, lo, hi) : solve(cdf, p, lo, hi);
    }

    /** 上側確率 q の点 (q が小さいときは q を直接解く)。 */
    static double invertRt(java.util.function.DoubleUnaryOperator cdf, java.util.function.DoubleUnaryOperator sf, double q, double lo, double hi) {
        return q < 0.5 ? solveDecreasing(sf, q, lo, hi) : solve(cdf, 1 - q, lo, hi);
    }

    /** t 分布の上側確率 (t >= 0)。 */
    static double tSf(double t, double df) {
        return 0.5 * betI(df / (df + t * t), df / 2, 0.5);
    }

    /** 自由度 df の t 分布の、上側確率 q (0 < q <= 0.5) の点。 */
    static double tInvUpper(double q, double df) {
        if (df > 5000) { // 自由度が大きいときは、不完全ベータの精度が足りないので正規分布からの展開 (Abramowitz-Stegun 26.7.5)
            double z = -normInv(q), z2 = z * z;
            double g1 = z * (z2 + 1) / 4;
            double g2 = z * (5 * z2 * z2 + 16 * z2 + 3) / 96;
            double g3 = z * (3 * z2 * z2 * z2 + 19 * z2 * z2 + 17 * z2 - 15) / 384;
            double g4 = z * (79 * z2 * z2 * z2 * z2 + 776 * z2 * z2 * z2 + 1482 * z2 * z2 - 1920 * z2 - 945) / 92160;
            return z + g1 / df + g2 / (df * df) + g3 / (df * df * df) + g4 / (df * df * df * df);
        }
        return solveDecreasing(t -> tSf(t, df), q, 0, 6);
    }

    /** 自由度 df の t 分布の、両側確率 p (0 < p < 2) の点 (p が 1 を超えると負)。 */
    static double tInv2T(double p, double df) {
        return p <= 1 ? tInvUpper(p / 2, df) : -tInvUpper((2 - p) / 2, df);
    }

    /** 下側確率 p の点 (0 < p < 1)。 */
    static double tInv(double p, double df) {
        if (p == 0.5) return 0;
        return p > 0.5 ? tInvUpper(1 - p, df) : -tInvUpper(p, df);
    }

    /** χ2 分布の CDF・上側確率・逆関数。 */
    static double chiCdf(double x, double df) {
        return x <= 0 ? 0 : gamP(df / 2.0, x / 2.0);
    }

    static double chiRt(double x, double df) {
        return x <= 0 ? 1 : gamQ(df / 2.0, x / 2.0);
    }

    static double chiInv(double p, double df) {
        return invert(x -> chiCdf(x, df), x -> chiRt(x, df), p, 0, Math.max(2, df + 4));
    }

    static double chiInvRt(double q, double df) {
        return invertRt(x -> chiCdf(x, df), x -> chiRt(x, df), q, 0, Math.max(2, df + 4));
    }

    static double chiPdf(double x, double df) {
        if (x < 0) return 0;
        double a = df / 2.0;
        return x == 0 ? (a < 1 ? Double.POSITIVE_INFINITY : a == 1 ? Math.exp(-logGamma(a)) / 2 : 0)
                : Math.exp((a - 1) * Math.log(x / 2) - x / 2 - logGamma(a)) / 2;
    }

    /** F 分布の CDF・逆関数・確率密度。 */
    static double fCdf(double x, double d1, double d2) {
        return x <= 0 ? 0 : betI(d1 * x / (d1 * x + d2), d1 / 2.0, d2 / 2.0);
    }

    /** F 分布の上側確率。 */
    static double fSf(double x, double d1, double d2) {
        return x <= 0 ? 1 : betI(d2 / (d2 + d1 * x), d2 / 2.0, d1 / 2.0);
    }

    static double fInv(double p, double d1, double d2) {
        return invert(x -> fCdf(x, d1, d2), x -> fSf(x, d1, d2), p, 0, 4);
    }

    static double fInvRt(double q, double d1, double d2) {
        return invertRt(x -> fCdf(x, d1, d2), x -> fSf(x, d1, d2), q, 0, 4);
    }

    static double fPdf(double x, double d1, double d2) {
        if (x <= 0) return 0;
        double a = d1 / 2.0, b = d2 / 2.0, z = a * x / (a * x + b);
        return Math.exp(a * Math.log(z) + b * Math.log1p(-z) - logBetaAB(a, b)) / x;
    }

    /** 二項係数を対数で (大きな trials でも壊れにくい)。 */
    static double logChoose(int n, int k) {
        return logGamma(n + 1d) - logGamma(k + 1d) - logGamma(n - k + 1d);
    }

    /** 二項分布の確率質量。 */
    static double binomPmf(int k, int n, double p) {
        if (k < 0 || k > n) return 0;
        if (p == 0) return k == 0 ? 1 : 0;
        if (p == 1) return k == n ? 1 : 0;
        return chooseSmall(n, k) * Math.pow(p, k) * Math.pow(1 - p, n - k);
    }

    /** 二項係数。n が小さい間は逐乗算 (0.375 のような値を 15 桁ぴったり出すため。大きい n は対数で)。 */
    static double chooseSmall(int n, int k) {
        if (k < 0 || k > n) return 0;
        int kk = Math.min(k, n - k);
        if (n > 1000) return Math.exp(logChoose(n, kk));
        double r = 1;
        for (int i = 1; i <= kk; i++) r = r * (n - kk + i) / i;
        return r;
    }


}
