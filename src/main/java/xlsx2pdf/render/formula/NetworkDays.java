package xlsx2pdf.render.formula;

import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.AggregateFunction.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** NETWORKDAYS.INTL (週末を指定できる稼働日数)。 */
final class NetworkDays {
    private NetworkDays() {
    }

    // ---- NETWORKDAYS.INTL ----

    /** 稼働日数 (始・終の日を含む)。weekend は 1〜17 の番号か 7 文字の文字列 (先頭が月曜、1 が休み)。 */
    static ValueEval networkDaysIntl(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 2 || a.length > 4) return ErrorEval.VALUE_INVALID;
        long s = dateOnly(a[0], ec), e = dateOnly(a[1], ec);
        if (s < 0 || e < 0) return ErrorEval.NUM_ERROR;
        java.util.BitSet wk = new java.util.BitSet(7); // 0=月曜 … 6=日曜
        if (omitted(a, 2)) {
            wk.set(5);
            wk.set(6); // 既定は土・日
        } else if (a[2] instanceof StringEval se) {
            String t = se.getStringValue();
            if (!t.matches("[01]{7}")) return ErrorEval.VALUE_INVALID; // 7 文字の 0/1 でない文字列
            for (int i = 0; i < 7; i++) if (t.charAt(i) == '1') wk.set(i);
        } else {
            long c = (long) num(a[2], ec);
            if (c < 1 || c > 17 || (c > 7 && c < 11)) return ErrorEval.NUM_ERROR; // 8・9・10 は無い
            if (c <= 7) {
                wk.set((int) ((c + 4) % 7)); // 1: 土・日, 2: 日・月, … 7: 金・土
                wk.set((int) ((c + 5) % 7));
            } else
                wk.set((int) ((c + 2) % 7)); // 11: 日, 12: 月, … 17: 土
        }
        java.util.TreeSet<Long> hol = new java.util.TreeSet<>();
        if (!omitted(a, 3) && !collectDays(a[3], hol)) return ErrorEval.VALUE_INVALID; // 休日列に日付として読めないセル (文字列など) があると Excel は #VALUE!
        // 始 > 終 のときは負の数。実 Excel は対称に両端を含む (NETWORKDAYS.INTL(1/10,1/1,1) = -8、
        // 1/1 から 1/10 の平日数 8 とちょうど符号だけ違う)
        return new NumberEval(s <= e ? workdays(s, e, wk, hol) : -workdays(e, s, wk, hol));
    }

    /** 日付の引数 (文字列の日付も読む) をシリアル値の整数部にする。 */
    static long dateOnly(ValueEval v, OperationEvaluationContext ec) throws EvaluationException {
        return (long) Math.floor(DateFunctions.dateArg(v, ec));
    }

    /** 曜日の番号 (0=月曜 … 6=日曜)。1900 暦の Excel は 60 番 (1900/2/29) を 61 番と同じ曜日にする。 */
    static int weekday(long n) {
        long k = FormulaContext.is1904() ? 5 : n <= 60 ? 6 : 5; // 1904/1/1 (土) = 0, 1900/1/1 (月) = 0 のずらした形
        return (int) ((n + k) % 7);
    }

    /** [lo, hi] の平日数 (両端を含む)。祝日と週末を数えない。 */
    static long workdays(long lo, long hi, java.util.BitSet wk, java.util.TreeSet<Long> hol) {
        if (lo > hi) return 0;
        if (!FormulaContext.is1904() && lo <= 60 && hi >= 61) // 60 番で曜日が 1 日ずれるので分けて数える
            return workdays(lo, 60, wk, hol) + workdays(61, hi, wk, hol);
        long n = hi - lo + 1, rem = n % 7;
        long off = (n / 7) * (long) wk.cardinality();
        for (int i = 0; i < rem; i++) if (wk.get(weekday(lo + i))) off++;
        long h = 0;
        for (long d : hol.subSet(lo, hi + 1)) if (!wk.get(weekday(d))) h++;
        return n - off - h;
    }

    /**
     * 休日列から日付の集合を読み出す。戻り値が false のセルが 1 つでもあったことを意味する
     * (実 Excel は休日列に文字列などの読むできないセルがあると #VALUE! を返す)。
     * 空セルは数えない。セル自体がエラーのときはそのエラーをそのまま伝える (Excel と同じ)。
     */
    static boolean collectDays(ValueEval v, java.util.NavigableSet<Long> out) throws EvaluationException {
        if (v instanceof ErrorEval e) throw new EvaluationException(e);
        else if (v instanceof AreaEval ar) {
            boolean ok = true;
            for (int r = 0; r < ar.getHeight(); r++)
                for (int c = 0; c < ar.getWidth(); c++)
                    ok &= addDay(ar.getRelativeValue(r, c), out);
            return ok;
        } else if (v instanceof RefEval re)
            return addDay(re.getInnerValueEval(re.getFirstSheetIndex()), out);
        else if (v instanceof org.apache.poi.ss.formula.eval.RefListEval rl) {
            boolean ok = true;
            for (Object x : rl.getList())
                ok &= collectDays((ValueEval) x, out);
            return ok;
        } else
            return addDay(v, out);
    }

    static boolean addDay(ValueEval v, java.util.NavigableSet<Long> out) {
        if (v == BlankEval.instance || v == org.apache.poi.ss.formula.eval.MissingArgEval.instance) return true;
        try {
            out.add((long) Math.floor(DateFunctions.dateArg(v)));
            return true;
        } catch (EvaluationException e) {
            return false; // 日付として読めないセル
        }
    }


}
