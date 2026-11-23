package xlsx2pdf.render.formula;

import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.function.FunctionMetadataRegistry;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.Function;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import xlsx2pdf.render.formula.FnSupport.Fn;

/**
 * POI にない (または Excel と結果が違う) 関数の登録。
 * 登録先は 2 か所ある: ブックごとの UDF の表 (POI が名前を知らない関数) と、POI の関数表 (JVM 全体で一度だけ。
 * 組み込みの枠や Analysis Toolpak の名前は UDF より先に見られるので、そちらにも置く)。
 */
final class FunctionRegistry {
    private FunctionRegistry() {
    }

    private static boolean globalDone;

    /** POI の実装が Excel と違う (6 引数の IPMT・PPMT、引数省略の ROMAN・DELTA など) ので、組み込みの表を書き換える関数。 */
    private static final java.util.Set<String> OVERRIDE_BUILTIN = java.util.Set.of("IPMT", "PPMT", "ROMAN", "DELTA", "SUBTOTAL");

    private static void replaceBuiltin(String name, Function f) {
        try {
            var meta = FunctionMetadataRegistry.getFunctionByName(name);
            if (meta == null) return;
            var field = FunctionEval.class.getDeclaredField("functions");
            field.setAccessible(true);
            ((Function[]) field.get(null))[meta.getIndex()] = f;
        } catch (ReflectiveOperationException | RuntimeException e) {
            // 書き換えられなければ POI の実装のまま
        }
    }

    /** 関数を登録する。ブックごとの UDF の表と、(一度だけ) POI の関数表へ。 */
    static synchronized void register(XSSFWorkbook wb) {
        if (!globalDone) installGlobal();
        Map<String, FreeRefFunction> m = new LinkedHashMap<>();
        define(m);
        wb.addToolPack(new DefaultUDFFinder(m.keySet().toArray(new String[0]), m.values().toArray(new FreeRefFunction[0])));
        globalDone = true;
    }

    /** POI の関数表への一度だけの登録 (組み込みの枠に入れる関数と、置き換える関数)。 */
    private static void installGlobal() {
        Date1904.install();
        TextFunctions.install();
        // DATEDIF と CELL は POI の組み込みの表に未実装の枠がある
        try {
            FunctionEval.registerFunction("DATEDIF", (args, row, col) -> {
                try {
                    return DateFunctions.datedif(args, row, col);
                } catch (EvaluationException e) {
                    return e.getErrorEval();
                }
            });
            FunctionEval.registerFunction("CELL", new CellFunction.CellFn());
        } catch (RuntimeException e) {
            // すでに登録されている
        }
    }

    private static void define(Map<String, FreeRefFunction> m) {
        // ブックごとの UDF の表だけに置く関数 (POI が名前を知らない新しい関数)
        udf(m, "XOR", InfoFunctions::xor);
        udf(m, "ISOWEEKNUM", DateFunctions::isoWeekNum);
        udf(m, "TEXTBEFORE", (a, ec) -> TextFunctions.textSplit(a, ec, true));
        udf(m, "TEXTAFTER", (a, ec) -> TextFunctions.textSplit(a, ec, false));
        // 統計 (順位・四分位・最頻)
        both(m, "RANK.EQ", (a, ec) -> StatFunctions.rank(a, ec, false));
        both(m, "RANK.AVG", (a, ec) -> StatFunctions.rank(a, ec, true));
        both(m, "MODE.SNGL", StatFunctions::modeSngl);
        both(m, "PERCENTILE.INC", (a, ec) -> StatFunctions.percentile(a, ec, true, false));
        both(m, "PERCENTILE.EXC", (a, ec) -> StatFunctions.percentile(a, ec, false, false));
        both(m, "QUARTILE.INC", (a, ec) -> StatFunctions.percentile(a, ec, true, true));
        both(m, "QUARTILE.EXC", (a, ec) -> StatFunctions.percentile(a, ec, false, true));
        // 稼働日数 (週末の指定)
        both(m, "NETWORKDAYS.INTL", NetworkDays::networkDaysIntl);
        // 数学
        both(m, "ERF", (a, ec) -> MathFunctions.erfFn(a, ec, 0));
        both(m, "ERF.PRECISE", (a, ec) -> MathFunctions.erfFn(a, ec, 1));
        both(m, "ERFC", (a, ec) -> MathFunctions.erfFn(a, ec, 2));
        both(m, "ERFC.PRECISE", (a, ec) -> MathFunctions.erfFn(a, ec, 2));
        both(m, "GAMMA", MathFunctions::gammaFn);
        both(m, "GAMMALN", MathFunctions::gammaLn);
        both(m, "PERMUT", MathFunctions::permut);
        both(m, "PERMUTATIONA", MathFunctions::permuta);
        both(m, "BITAND", (a, ec) -> MathFunctions.bitwise(a, ec, 0));
        both(m, "BITOR", (a, ec) -> MathFunctions.bitwise(a, ec, 1));
        both(m, "BITXOR", (a, ec) -> MathFunctions.bitwise(a, ec, 2));
        both(m, "BITLSHIFT", (a, ec) -> MathFunctions.bitShift(a, ec, true));
        both(m, "BITRSHIFT", (a, ec) -> MathFunctions.bitShift(a, ec, false));
        both(m, "ISPMT", FinanceFunctions::ispmt);
        // 情報の類
        both(m, "ISFORMULA", InfoFunctions::isFormula);
        both(m, "N", InfoFunctions::n);
        both(m, "TYPE", InfoFunctions::type);
        // ---- 統計 (*.DIST / *.INV と旧名) ----
        both(m, "BETA.DIST", (a, ec) -> DistFunctions.betaDist(a, ec, false)); // 新名は cumulative 必須
        both(m, "BETADIST", (a, ec) -> DistFunctions.betaDist(a, ec, true)); // 旧名は cumulative 省略可 (3 引数でも通る)
        both(m, "BETA.INV", (a, ec) -> DistFunctions.betaInv(a, ec, false));
        both(m, "BETAINV", (a, ec) -> DistFunctions.betaInv(a, ec, false));
        both(m, "BINOM.DIST", (a, ec) -> DistFunctions.binomDist(a, ec, true));
        both(m, "BINOMDIST", (a, ec) -> DistFunctions.binomDist(a, ec, true));
        both(m, "BINOM.DIST.RANGE", DistFunctions::binomRange);
        both(m, "CRITBINOM", DistFunctions::critBinom);
        both(m, "CHISQ.DIST", (a, ec) -> DistFunctions.chi2Dist(a, ec, true));
        both(m, "CHISQ.DIST.RT", (a, ec) -> DistFunctions.chi2Dist(a, ec, false));
        both(m, "CHIDIST", (a, ec) -> DistFunctions.chi2Dist(a, ec, false));
        both(m, "CHISQ.INV", (a, ec) -> DistFunctions.chi2Inv(a, ec, false));
        both(m, "CHISQ.INV.RT", (a, ec) -> DistFunctions.chi2Inv(a, ec, true));
        both(m, "CHIINV", (a, ec) -> DistFunctions.chi2Inv(a, ec, true));
        both(m, "CHISQ.TEST", DistFunctions::chi2Test);
        both(m, "CHITEST", DistFunctions::chi2Test);
        both(m, "F.DIST", (a, ec) -> DistFunctions.fDist(a, ec, true));
        both(m, "F.DIST.RT", (a, ec) -> DistFunctions.fDist(a, ec, false));
        both(m, "FDIST", (a, ec) -> DistFunctions.fDist(a, ec, false));
        both(m, "F.INV", (a, ec) -> DistFunctions.fInvFn(a, ec, false));
        both(m, "F.INV.RT", (a, ec) -> DistFunctions.fInvFn(a, ec, true));
        both(m, "FINV", (a, ec) -> DistFunctions.fInvFn(a, ec, true));
        both(m, "F.TEST", DistFunctions::fTest);
        both(m, "FTEST", DistFunctions::fTest);
        both(m, "GAMMA.DIST", (a, ec) -> DistFunctions.gammaDist(a, ec, true));
        both(m, "GAMMADIST", (a, ec) -> DistFunctions.gammaDist(a, ec, true));
        both(m, "GAMMA.INV", DistFunctions::gammaInv);
        both(m, "GAMMAINV", DistFunctions::gammaInv);
        both(m, "GAMMALN.PRECISE", MathFunctions::gammaLn);
        both(m, "LOGNORM.DIST", (a, ec) -> DistFunctions.logNormDist(a, ec, true));
        both(m, "LOGNORMDIST", (a, ec) -> DistFunctions.logNormDist(a, ec, false));
        both(m, "LOGNORM.INV", DistFunctions::logNormInv);
        both(m, "NORM.S.DIST", DistFunctions::normSDist);
        both(m, "NORM.S.INV", DistFunctions::normSInv);
        both(m, "NORM.DIST", DistFunctions::normDist);
        both(m, "NORM.INV", DistFunctions::normInvFn);
        both(m, "LOGINV", DistFunctions::logNormInv);
        both(m, "NEGBINOM.DIST", (a, ec) -> DistFunctions.negBinom(a, ec, true));
        both(m, "NEGBINOMDIST", (a, ec) -> DistFunctions.negBinom(a, ec, false));
        both(m, "HYPGEOM.DIST", (a, ec) -> DistFunctions.hyperGeom(a, ec, true));
        both(m, "HYPGEOMDIST", (a, ec) -> DistFunctions.hyperGeom(a, ec, false));
        both(m, "EXPON.DIST", (a, ec) -> DistFunctions.exponDist(a, ec, true));
        both(m, "EXPONDIST", (a, ec) -> DistFunctions.exponDist(a, ec, true)); // 旧名も 3 引数 (cumulative を持つ)
        both(m, "WEIBULL.DIST", (a, ec) -> DistFunctions.weibull(a, ec, true));
        both(m, "WEIBULL", (a, ec) -> DistFunctions.weibull(a, ec, true));
        both(m, "T.INV", (a, ec) -> DistFunctions.tInvFn(a, ec, true));
        both(m, "T.INV.2T", (a, ec) -> DistFunctions.tInvFn(a, ec, false));
        both(m, "TINV", (a, ec) -> DistFunctions.tInvFn(a, ec, false));
        both(m, "Z.TEST", DistFunctions::zTest);
        both(m, "ZTEST", DistFunctions::zTest);
        both(m, "CONFIDENCE", (a, ec) -> DistFunctions.confidence(a, ec, false));
        both(m, "CONFIDENCE.NORM", (a, ec) -> DistFunctions.confidence(a, ec, false));
        both(m, "CONFIDENCE.T", (a, ec) -> DistFunctions.confidence(a, ec, true));
        both(m, "RSQ", StatFunctions::rsq);
        both(m, "STEYX", StatFunctions::steyx);
        both(m, "KURT", StatFunctions::kurt);
        both(m, "SKEW", StatFunctions::skew);
        both(m, "HARMEAN", StatFunctions::harmean);
        both(m, "TRIMMEAN", StatFunctions::trimMean);
        both(m, "PROB", StatFunctions::prob);
        both(m, "FISHER", (a, ec) -> StatFunctions.fisher(a, ec, false));
        both(m, "FISHERINV", (a, ec) -> StatFunctions.fisher(a, ec, true));
        // ---- 減価償却 ----
        both(m, "SLN", FinanceFunctions::sln);
        both(m, "SYD", FinanceFunctions::syd);
        both(m, "DDB", FinanceFunctions::ddb);
        both(m, "DB", FinanceFunctions::db);
        both(m, "VDB", FinanceFunctions::vdb);
        // ---- POI に無い・枠だけの関数 ----
        both(m, "PHI", NumberFunctions::phi);
        both(m, "GESTEP", NumberFunctions::gestep);
        both(m, "BASE", NumberFunctions::base);
        both(m, "DECIMAL", NumberFunctions::decimal);
        both(m, "ARABIC", NumberFunctions::arabic);
        both(m, "UNICHAR", NumberFunctions::unichar);
        both(m, "UNICODE", NumberFunctions::unicode);
        both(m, "CONVERT", ConvertFunction::convert);
        both(m, "AGGREGATE", AggregateFunction::aggregate);
        both(m, "SUBTOTAL", AggregateFunction::subtotal);
        both(m, "XNPV", FinanceFunctions::xnpv);
        both(m, "XIRR", FinanceFunctions::xirr);
        both(m, "EFFECT", FinanceFunctions::effect);
        both(m, "NOMINAL", FinanceFunctions::nominal);
        both(m, "PDURATION", FinanceFunctions::pduration);
        both(m, "RRI", FinanceFunctions::rri);
        both(m, "CUMIPMT", FinanceFunctions::cumipmt);
        both(m, "CUMPRINC", FinanceFunctions::cumprinc);
        // POI 版の挙動が実測と違うものだけ上書き (IPMT の 1 期目 #VALUE!、DELTA の既定値、ROMAN の 1 引数)
        both(m, "DELTA", NumberFunctions::delta);
        both(m, "ROMAN", NumberFunctions::roman);
        both(m, "IPMT", FinanceFunctions::ipmt);
        both(m, "PPMT", FinanceFunctions::ppmt);
    }

    private static void udf(Map<String, FreeRefFunction> m, String name, Fn fn) {
        FreeRefFunction f = (args, ec) -> {
            try {
                return fn.apply(args, ec);
            } catch (EvaluationException e) {
                return e.getErrorEval();
            }
        };
        m.put(name, f);
        m.put("_xlfn." + name, f);
    }

    /** POI の組み込み関数の枠の形 (引数のセルの位置は FnSupport.POS に入れて渡す)。 */
    private static Function builtin(Fn fn) {
        return (args, row, col) -> {
            FnSupport.POS.set(new int[] {row, col});
            try {
                return fn.apply(args, null);
            } catch (EvaluationException e) {
                return e.getErrorEval();
            } finally {
                FnSupport.POS.remove();
            }
        };
    }

    /** UDF の表と、POI の組み込みの枠・Analysis Toolpak の両方に置く (POI の関数表への登録は一度だけ)。 */
    private static void both(Map<String, FreeRefFunction> m, String name, Fn fn) {
        udf(m, name, fn);
        if (globalDone) return;
        Function b = builtin(fn);
        for (String n : new String[] {name, "_xlfn." + name}) {
            try {
                FunctionEval.registerFunction(n, b);
            } catch (RuntimeException e) {
                // POI が実装済みの関数は registerFunction が拒否する。結果が Excel と違うものだけ関数表を直接書き換える
                if (OVERRIDE_BUILTIN.contains(name)) replaceBuiltin(n, b);
            }
            // POI が ATP で実装済みの名前 (NORM.S.DIST など) は、2 引数版では上書きを拒否される。3 引数版は上書きできる
            FreeRefFunction atp = (args, ec) -> b.evaluate(args, ec == null ? -1 : ec.getRowIndex(), ec == null ? -1 : ec.getColumnIndex());
            try {
                AnalysisToolPak.registerFunction(n, atp);
            } catch (RuntimeException e) {
                try {
                    AnalysisToolPak.registerFunction(n, atp, true);
                } catch (RuntimeException e2) {
                    // ATP の枠に無い名前は UDF 側で拾われる
                }
            }
        }
    }
}
