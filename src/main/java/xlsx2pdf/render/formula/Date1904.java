package xlsx2pdf.render.formula;

import java.util.function.Predicate;
import org.apache.poi.ss.formula.CacheAreaEval;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.function.FunctionMetadataRegistry;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.Function;

/**
 * 1904 年の日付システムのブック向けに、POI の日付の関数 (1900 年のシリアル値で計算する) を包む。
 * 日付の引数に 1462 日 (1900 年と 1904 年のシリアル値の差) を足してから POI の関数を呼び、日付を返す関数は結果から 1462 を引く。
 * 登録は JVM 全体で一度だけで、包んだ関数は呼ばれるたびに現在のブックが 1904 年かを見る。
 */
final class Date1904 {
    private Date1904() {
    }

    private static final double OFFSET = 1462;
    private static boolean installed;

    /** 引数の位置が日付か。 */
    private record Spec(String name, Predicate<Integer> dateArg, boolean dateResult) {
    }

    private static final Predicate<Integer> NONE = i -> false, FIRST = i -> i == 0;

    private static final Spec[] SPECS = {
        new Spec("DATE", NONE, true), new Spec("DATEVALUE", NONE, true),
        new Spec("YEAR", FIRST, false), new Spec("MONTH", FIRST, false), new Spec("DAY", FIRST, false),
        new Spec("WEEKDAY", FIRST, false), new Spec("WEEKNUM", FIRST, false),
        new Spec("EDATE", FIRST, true), new Spec("EOMONTH", FIRST, true),
        new Spec("WORKDAY", i -> i == 0 || i == 2, true), new Spec("NETWORKDAYS", i -> i <= 1 || i == 2, false),
        new Spec("WORKDAY.INTL", i -> i == 0 || i == 3, true),
        new Spec("DAYS", i -> i <= 1, false), new Spec("DAYS360", i -> i <= 1, false), new Spec("YEARFRAC", i -> i <= 1, false),
    };

    static synchronized void install() {
        if (installed) return;
        installed = true;
        for (Spec sp : SPECS) {
            var meta = FunctionMetadataRegistry.getFunctionByName(sp.name);
            if (meta != null) {
                try {
                    var field = FunctionEval.class.getDeclaredField("functions");
                    field.setAccessible(true);
                    Function[] table = (Function[]) field.get(null);
                    Function orig = table[meta.getIndex()];
                    if (orig != null) table[meta.getIndex()] = (args, row, col) -> d1904() ? run(sp, args, row, col, (a, r, c) -> orig.evaluate(a, r, c)) : orig.evaluate(args, row, col);
                } catch (ReflectiveOperationException | RuntimeException e) {
                    // 包めなければ 1900 年の計算のまま
                }
            } else {
                for (String n : new String[] {sp.name, "_xlfn." + sp.name}) {
                    FreeRefFunction orig = AnalysisToolPak.instance.findFunction(n);
                    if (orig == null) continue;
                    try {
                        AnalysisToolPak.registerFunction(n, (args, ec) -> d1904() ? runAtp(sp, args, ec, orig) : orig.evaluate(args, ec), true);
                    } catch (RuntimeException e) {
                        // 包めなければ 1900 年の計算のまま
                    }
                }
            }
        }
    }

    private interface Call {
        ValueEval call(ValueEval[] a, int row, int col);
    }

    private static boolean d1904() {
        return FormulaContext.is1904();
    }

    private static ValueEval run(Spec sp, ValueEval[] args, int row, int col, Call c) {
        ValueEval[] a = new ValueEval[args.length];
        for (int i = 0; i < a.length; i++) a[i] = sp.dateArg.test(i) ? shift(args[i]) : args[i];
        for (int i = 0; i < a.length; i++)
            if (sp.dateArg.test(i) && a[i] instanceof NumberEval n && n.getNumberValue() < OFFSET) return ErrorEval.NUM_ERROR;
        return out(sp, c.call(a, row, col));
    }

    private static ValueEval runAtp(Spec sp, ValueEval[] args, OperationEvaluationContext ec, FreeRefFunction orig) {
        ValueEval[] a = new ValueEval[args.length];
        for (int i = 0; i < a.length; i++) a[i] = sp.dateArg.test(i) ? shift(args[i]) : args[i];
        for (int i = 0; i < a.length; i++)
            if (sp.dateArg.test(i) && a[i] instanceof NumberEval n && n.getNumberValue() < OFFSET) return ErrorEval.NUM_ERROR;
        return out(sp, orig.evaluate(a, ec));
    }

    private static ValueEval out(Spec sp, ValueEval r) {
        if (!sp.dateResult || !(r instanceof NumberEval n)) return r;
        double v = n.getNumberValue() - OFFSET;
        if (v < 0) return sp.name.equals("DATEVALUE") ? ErrorEval.VALUE_INVALID : ErrorEval.NUM_ERROR;
        return new NumberEval(v);
    }

    /** 日付の引数 (数値・セル・範囲) を 1900 年のシリアル値にする。 */
    private static ValueEval shift(ValueEval v) {
        if (v instanceof NumberEval n) return new NumberEval(n.getNumberValue() + OFFSET);
        if (v instanceof RefEval re) {
            ValueEval iv = re.getInnerValueEval(re.getFirstSheetIndex());
            return iv instanceof NumberEval n ? new NumberEval(n.getNumberValue() + OFFSET) : iv;
        }
        if (v instanceof AreaEval ar) {
            ValueEval[] vals = new ValueEval[ar.getHeight() * ar.getWidth()];
            for (int r = 0; r < ar.getHeight(); r++)
                for (int c = 0; c < ar.getWidth(); c++) {
                    ValueEval x = ar.getRelativeValue(r, c);
                    vals[r * ar.getWidth() + c] = x instanceof NumberEval n ? new NumberEval(n.getNumberValue() + OFFSET) : x;
                }
            return new CacheAreaEval(ar.getFirstRow(), ar.getFirstColumn(), ar.getLastRow(), ar.getLastColumn(), vals);
        }
        return v;
    }
}
