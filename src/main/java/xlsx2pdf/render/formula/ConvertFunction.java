package xlsx2pdf.render.formula;

import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.NumberFunctions.isBool;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.ValueEval;

/**
 * CONVERT(数値, 元の単位, 先の単位)。単位の名前と換算の係数は Excel 2024 に実際に計算させて取り出した
 * (全単位の名前と接頭辞の組み合わせ 5,000 通りを、各種類の基準単位へ換算した)。
 *
 * <p>単位は種類ごとに基準単位 (長さ m・質量 g・時間 sec・圧力 Pa・力 N・エネルギー J・仕事率 W・磁束密度 T・体積 l・面積 m2・
 * 情報 bit・速さ m/s・温度 K) への係数で表す。換算は元の係数 ÷ 先の係数。温度だけは係数のほかにずれがある。
 * 種類が違う単位や、知らない単位は #N/A。
 */
final class ConvertFunction {
    private ConvertFunction() {
    }

    /** 単位の定義 (係数・接頭辞の種類)。 */
    private record Unit(String category, double factor, char prefixes) {
    }

    /** 単位と接頭辞を分けた結果 (接頭辞の倍率を含めた基準単位への係数)。 */
    private record Resolved(String category, double factor) {
    }

    private static final Map<String, Unit> UNITS = new HashMap<>();

    /** 温度 (基準 K) の「係数, ずれ」。K = 係数 × 値 + ずれ。 */
    private static final Map<String, double[]> TEMPERATURE = new HashMap<>();

    /** 接頭辞 (SI)。da と e はどちらも 10 倍。µ は使えず u。 */
    private static final String[] SI_PREFIX = {"Y", "Z", "E", "P", "T", "G", "M", "k", "h", "e", "da", "d", "c", "m", "u", "n", "p", "f", "a", "z", "y"};
    private static final double[] SI_VALUE = {1e24, 1e21, 1e18, 1e15, 1e12, 1e9, 1e6, 1e3, 1e2, 10, 10, 0.1, 0.01, 1e-3, 1e-6, 1e-9, 1e-12, 1e-15, 1e-18, 1e-21, 1e-24};
    /** 2 進接頭辞 (bit・byte だけ)。 */
    private static final String[] BIN_PREFIX = {"ki", "Mi", "Gi", "Ti", "Pi", "Ei", "Zi", "Yi"};

    /**
     * 単位の表: 「名前,種類,基準単位への係数,接頭辞」。接頭辞は N = つかない、M = SI、B = SI と 2 進 (bit・byte)、
     * S = SI の 2 乗 (面積: km2 は 1e6 m2)、C = SI の 3 乗 (体積)。
     */
    private static final String[] TABLE = {
        "m,DIST,1.0,M",
        "mi,DIST,1609.344,N",
        "Nmi,DIST,1852.0,N",
        "in,DIST,0.0254,N",
        "ft,DIST,0.3048,N",
        "yd,DIST,0.9144,N",
        "ang,DIST,1e-10,M",
        "ell,DIST,1.143,N",
        "ly,DIST,9460730472580800.0,M",
        "parsec,DIST,3.0856775812815532e+16,M",
        "pc,DIST,3.0856775812815532e+16,M",
        "Pica,DIST,0.0003527777777777778,N",
        "pica,DIST,0.004233333333333334,N",
        "g,MASS,1.0,M",
        "sg,MASS,14593.902937206363,N",
        "lbm,MASS,453.59237,N",
        "u,MASS,1.660538782e-24,M",
        "ozm,MASS,28.349523125,N",
        "grain,MASS,0.06479891,N",
        "cwt,MASS,45359.237,N",
        "shweight,MASS,45359.237,N",
        "uk_cwt,MASS,50802.34544,N",
        "lcwt,MASS,50802.34544,N",
        "hweight,MASS,50802.34544,N",
        "stone,MASS,6350.29318,N",
        "ton,MASS,907184.74,N",
        "uk_ton,MASS,1016046.9088,N",
        "LTON,MASS,1016046.9088,N",
        "brton,MASS,1016046.9088,N",
        "yr,TIME,31557600.0,N",
        "day,TIME,86400.0,N",
        "d,TIME,86400.0,N",
        "hr,TIME,3600.0,N",
        "mn,TIME,60.0,N",
        "min,TIME,60.0,N",
        "sec,TIME,1.0,M",
        "s,TIME,1.0,M",
        "Pa,PRES,1.0,M",
        "p,PRES,1.0,M",
        "atm,PRES,101325.0,M",
        "at,PRES,101325.0,M",
        "mmHg,PRES,133.322,M",
        "psi,PRES,6894.757293168362,N",
        "Torr,PRES,133.32236842105263,N",
        "N,FORCE,1.0,M",
        "dyn,FORCE,1e-05,M",
        "dy,FORCE,1e-05,M",
        "lbf,FORCE,4.4482216152605005,N",
        "pond,FORCE,0.00980665,M",
        "J,ENERGY,1.0,M",
        "e,ENERGY,1e-07,M",
        "c,ENERGY,4.184,M",
        "cal,ENERGY,4.1868,M",
        "eV,ENERGY,1.602176487e-19,M",
        "ev,ENERGY,1.602176487e-19,M",
        "HPh,ENERGY,2684519.537696173,N",
        "hh,ENERGY,2684519.537696173,N",
        "Wh,ENERGY,3600.0,M",
        "wh,ENERGY,3600.0,M",
        "flb,ENERGY,1.3558179483314003,N",
        "BTU,ENERGY,1055.05585262,N",
        "btu,ENERGY,1055.05585262,N",
        "HP,POWER,745.6998715822702,N",
        "h,POWER,745.6998715822702,N",
        "PS,POWER,735.49875,N",
        "W,POWER,1.0,M",
        "w,POWER,1.0,M",
        "T,MAG,1.0,M",
        "ga,MAG,0.0001,M",
        "tsp,VOL,0.00492892159375,N",
        "tspm,VOL,0.005,N",
        "tbs,VOL,0.01478676478125,N",
        "oz,VOL,0.0295735295625,N",
        "cup,VOL,0.2365882365,N",
        "pt,VOL,0.473176473,N",
        "us_pt,VOL,0.473176473,N",
        "uk_pt,VOL,0.56826125,M",
        "qt,VOL,0.946352946,N",
        "uk_qt,VOL,1.1365225,N",
        "gal,VOL,3.785411784,N",
        "uk_gal,VOL,4.54609,N",
        "l,VOL,1.0,M",
        "L,VOL,1.0,M",
        "lt,VOL,1.0,M",
        "m3,VOL,1000.0,C",
        "mi3,VOL,4168181825440.5796,N",
        "Nmi3,VOL,6352182208000.0,N",
        "in3,VOL,0.016387064,N",
        "ft3,VOL,28.316846592,N",
        "yd3,VOL,764.554857984,N",
        "ang3,VOL,1e-27,C",
        "Pica3,VOL,4.39039566186557e-08,N",
        "barrel,VOL,158.987294928,N",
        "bushel,VOL,35.23907016688,N",
        "regton,VOL,2831.6846592,N",
        "GRT,VOL,2831.6846592,N",
        "MTON,VOL,1132.67386368,N",
        "ly3,VOL,8.467866646237152e+50,N",
        "uk_acre,AREA,4046.8564224,N",
        "us_acre,AREA,4046.8726098742522,N",
        "ang2,AREA,1e-20,S",
        "ar,AREA,100.0,M",
        "ft2,AREA,0.09290304,N",
        "ha,AREA,10000.0,N",
        "in2,AREA,0.00064516,N",
        "ly2,AREA,8.95054210748189e+31,N",
        "m2,AREA,1.0,S",
        "Morgen,AREA,2500.0,N",
        "mi2,AREA,2589988.110336,N",
        "Nmi2,AREA,3429904.0,N",
        "Pica2,AREA,1.2445216049382715e-07,N",
        "yd2,AREA,0.83612736,N",
        "bit,INFO,1.0,B",
        "byte,INFO,8.0,B",
        "m/s,SPEED,1.0,M",
        "m/sec,SPEED,1.0,M",
        "m/h,SPEED,0.0002777777777777778,M",
        "m/hr,SPEED,0.0002777777777777778,M",
        "mph,SPEED,0.44704,M",
        "kn,SPEED,0.5144444444444445,N",
        "admkn,SPEED,0.5147733333333333,N",
    };

    static {
        for (String line : TABLE) {
            String[] p = line.split(",");
            UNITS.put(p[0], new Unit(p[1], Double.parseDouble(p[2]), p[3].charAt(0)));
        }
        TEMPERATURE.put("K", new double[] {1, 0});
        TEMPERATURE.put("kel", new double[] {1, 0});
        TEMPERATURE.put("C", new double[] {1, 273.15});
        TEMPERATURE.put("cel", new double[] {1, 273.15});
        TEMPERATURE.put("F", new double[] {5 / 9d, 459.67 * 5 / 9d});
        TEMPERATURE.put("fah", new double[] {5 / 9d, 459.67 * 5 / 9d});
        TEMPERATURE.put("Rank", new double[] {5 / 9d, 0});
        TEMPERATURE.put("Reau", new double[] {1.25, 273.15});
    }

    /** 単位の名前を、(接頭辞 + 素の単位) の分け方ごとに解く。分け方が複数あるものがある (pc は パーセク と ピコ + c)。 */
    private static List<Resolved> resolve(String s) {
        List<Resolved> out = new ArrayList<>();
        Unit u = UNITS.get(s);
        if (u != null) out.add(new Resolved(u.category, u.factor));
        if (TEMPERATURE.containsKey(s)) out.add(new Resolved("TEMP", 0));
        for (int i = 0; i < SI_PREFIX.length; i++) {
            if (!s.startsWith(SI_PREFIX[i])) continue;
            String rest = s.substring(SI_PREFIX[i].length());
            double m = SI_VALUE[i];
            Unit r = UNITS.get(rest);
            if (r != null && (r.prefixes == 'M' || r.prefixes == 'B' || r.prefixes == 'S' || r.prefixes == 'C'))
                out.add(new Resolved(r.category, r.factor * (r.prefixes == 'S' ? m * m : r.prefixes == 'C' ? m * m * m : m)));
            if (rest.equals("K") || rest.equals("kel")) out.add(new Resolved("TEMP", 0)); // mK・kK など (温度は K だけ接頭辞がつく)
        }
        for (int i = 0; i < BIN_PREFIX.length && s.length() > 2; i++) {
            if (!s.startsWith(BIN_PREFIX[i])) continue;
            Unit r = UNITS.get(s.substring(2));
            if (r != null && r.prefixes == 'B') out.add(new Resolved(r.category, r.factor * Math.pow(2, 10 * (i + 1))));
        }
        return out;
    }

    static ValueEval convert(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 3)
            return ErrorEval.VALUE_INVALID;
        // 論理値は #VALUE!
        if (isBool(a[0], ec) || isBool(a[1], ec) || isBool(a[2], ec))
            return ErrorEval.VALUE_INVALID;
        double v = num(a[0], ec);
        String from = str(a[1], ec), to = str(a[2], ec);
        for (Resolved f : resolve(from)) {
            for (Resolved t : resolve(to)) {
                if (!f.category.equals(t.category)) continue;
                double r;
                if (f.category.equals("TEMP")) {
                    double[] x = temperature(from), y = temperature(to);
                    if (x == null || y == null) continue;
                    // K に接頭辞がつく (mK・kK など) ときの倍率は、相手が Reau だと無視され、Rank へは元の側だけ無視される (Excel の結果)
                    boolean ignoreX = to.equals("Rank") || to.equals("Reau"), ignoreY = from.equals("Reau");
                    double kelvin = v * x[0] * (ignoreX ? 1 : x[2]) + x[1];
                    r = (kelvin - y[1]) / (y[0] * (ignoreY ? 1 : y[2]));                } else {
                    r = v * (f.factor / t.factor);
                }
                return Double.isNaN(r) || Double.isInfinite(r) ? ErrorEval.NUM_ERROR : new NumberEval(r);
            }
        }
        return ErrorEval.NA;
    }

    /** 温度の単位 (接頭辞つきの mK・kK を含む) の {係数, ずれ, 接頭辞の倍率}。K = 係数 × 倍率 × 値 + ずれ。 */
    private static double[] temperature(String s) {
        double[] t = TEMPERATURE.get(s);
        if (t != null) return new double[] {t[0], t[1], 1};
        for (int i = 0; i < SI_PREFIX.length; i++) {
            if (!s.startsWith(SI_PREFIX[i])) continue;
            String rest = s.substring(SI_PREFIX[i].length());
            if (rest.equals("K") || rest.equals("kel")) return new double[] {1, 0, SI_VALUE[i]};
        }
        return null;
    }
}
