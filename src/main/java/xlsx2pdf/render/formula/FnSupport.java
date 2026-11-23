package xlsx2pdf.render.formula;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** 数式関数の実装が共通で使う引数の読み取り・数値の収集。 */
final class FnSupport {
    private FnSupport() {
    }

    /** 関数の実装の形 (評価できないときは EvaluationException でエラー値を返す)。 */
    interface Fn {
        ValueEval apply(ValueEval[] args, OperationEvaluationContext ec) throws EvaluationException;
    }

    // ---- 引数の読み取り ----

    /** POI の組み込み関数の枠を置き換える道 (`FunctionEval.registerFunction`) では、引数のセルの位置がこの箱に入ってくる。 */
    static final ThreadLocal<int[]> POS = new ThreadLocal<>();

    /** 1 つの関数の中で回すループの回数の上限 (極端な引数で変換が終わらなくならないように)。 */
    static final long MAX_STEPS = 10_000_000;

    /**
     * 回数 n のループを回してよいか確かめる。超えると例外にする (--recalc はそのセルを「計算できない数式」として、保存された結果を使う)。
     */
    static void checkSteps(double n) {
        if (!(n <= MAX_STEPS)) throw new IllegalArgumentException("計算の回数が多すぎます (" + n + ")");
    }

    static int rowOf(OperationEvaluationContext ec) {
        int[] p = POS.get();
        return ec != null ? ec.getRowIndex() : p == null ? -1 : p[0];
    }

    static int colOf(OperationEvaluationContext ec) {
        int[] p = POS.get();
        return ec != null ? ec.getColumnIndex() : p == null ? -1 : p[1];
    }

    static ValueEval scalar(ValueEval v, OperationEvaluationContext ec) throws EvaluationException {
        return OperandResolver.getSingleValue(v, rowOf(ec), colOf(ec));
    }

    static double num(ValueEval v, OperationEvaluationContext ec) throws EvaluationException {
        return OperandResolver.coerceValueToDouble(scalar(v, ec));
    }

    static String str(ValueEval v, OperationEvaluationContext ec) throws EvaluationException {
        return OperandResolver.coerceValueToString(scalar(v, ec));
    }

    static boolean omitted(ValueEval[] a, int i) {
        return a.length <= i || a[i] == BlankEval.instance || a[i] == org.apache.poi.ss.formula.eval.MissingArgEval.instance;
    }


    // ---- 数値の収集 (統計関数) ----

    /** 文字列を数値に読む (読めなければ null)。 */
    static Double numberOrNull(String s) {
        try {
            return Double.valueOf(s.trim());
        } catch (NumberFormatException e) {
            return null;
        }
    }

    /** 統計関数の引数から数値だけを列にする。 */
    static List<Double> numbers(ValueEval v) throws EvaluationException {
        List<Double> out = new ArrayList<>();
        collect(v, false, out);
        return out;
    }

    /**
     * Excel の統計関数と同じ集め方: セル参照の中の数値だけを使う (参照の中の文字列・論理値・空は数えない)。
     * 直接書いた文字列は #VALUE! 、どの場所でもエラー値はそのまま伝える。
     */
    static void collect(ValueEval v, boolean fromRef, List<Double> out) throws EvaluationException {
        if (v instanceof ErrorEval e) throw new EvaluationException(e);
        else if (v instanceof AreaEval ar) {
            for (int r = 0; r < ar.getHeight(); r++)
                for (int c = 0; c < ar.getWidth(); c++) collect(ar.getRelativeValue(r, c), true, out);
        } else if (v instanceof RefEval re)
            collect(re.getInnerValueEval(re.getFirstSheetIndex()), true, out);
        else if (v instanceof org.apache.poi.ss.formula.eval.RefListEval rl) {
            for (Object x : rl.getList()) collect((ValueEval) x, true, out);
        } else if (v == BlankEval.instance || v == org.apache.poi.ss.formula.eval.MissingArgEval.instance) {
            // 空は無視
        } else if (fromRef && !(v instanceof NumberEval)) {
            // 参照の中の文字列・論理値は無視 (Excel)
        } else if (v instanceof NumberEval n) out.add(n.getNumberValue());
        else if (v instanceof BoolEval b) out.add(b.getBooleanValue() ? 1d : 0d);
        else out.add(OperandResolver.coerceValueToDouble(v));
    }


}
