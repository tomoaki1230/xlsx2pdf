package xlsx2pdf.render.formula;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.InfoFunctions.*;

/** AGGREGATE (関数番号とオプションで集計)。 */
final class AggregateFunction {
    private AggregateFunction() {
    }

    // ---- AGGREGATE: function_num・options で集計 (隠し行・エラーを無視できる) ----

    static ValueEval aggregate(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 3)
            return ErrorEval.VALUE_INVALID;
        double fd = num(a[0], ec), od = num(a[1], ec);
        int fn = (int) Math.floor(fd), opt = (int) Math.floor(od);
        // 実測: function_num が 1..19 を外れる (0・-1・20・21・22..26・100) と #VALUE! (#NUM! ではない)。
        //       20・21 は 3 引数形 (`AGGREGATE(20,0,A211:G211)`) でも #VALUE! なので、有効な番号そのものではない。
        //       options も 0..7 を外れる (8..13・14..21・22・-1・100) と #VALUE!。14..21 は単一範囲では使えない。
        if (fn < 1 || fn > 19 || opt < 0 || opt > 7)
            return ErrorEval.VALUE_INVALID;
        // options: 1・3・5・7 は隠し行を無視、2・3・6・7 はエラー値を無視、0〜3 は入れ子の SUBTOTAL・AGGREGATE を無視 (Excel の説明どおり。隠し列は無視しない)
        boolean propagate = fn != 2 && fn != 3;
        Opts opts = new Opts(opt == 1 || opt == 3 || opt == 5 || opt == 7, opt == 2 || opt == 3 || opt == 6 || opt == 7, opt <= 3, propagate);
        // 実測: 4 番目以降の引数 (k・q) を受け付けるのは 14..19 (LARGE・SMALL・PERCENTILE.INC・QUARTILE.INC・
        //       PERCENTILE.EXC・QUARTILE.EXC) だけ。それ以外に 4 番目を渡すと #VALUE! (実測 AGGREGATE(9,0,A207:C207,1))。
        boolean kFn = fn >= 14 && fn <= 19;
        if (!aggRefsOk(a, kFn))
            return ErrorEval.VALUE_INVALID;
        List<Double> nums = new ArrayList<>();
        int[] cnt = new int[2]; // [数値の個数 (COUNT), 空でない個数 (COUNTA)]
        int dataEnd = kFn ? 3 : a.length; // k を取る関数では 4 番目はデータではない
        for (int i = 2; i < dataEnd; i++)
            aggCollect(a[i], true, opts, nums, cnt);
        java.util.Collections.sort(nums);
        ValueEval basic = summary(fn, nums, cnt);
        if (basic != null)
            return basic;
        switch (fn) {
            case 16: return percentileFrom(nums, k(a, 3, ec), false);
            case 18: return percentileFrom(nums, k(a, 3, ec), true);
            case 17: return percentileFrom(nums, (long) k(a, 3, ec) / 4.0, false);
            case 19: return percentileFrom(nums, (long) k(a, 3, ec) / 4.0, true);
            case 12: return statsMedian(nums);
            case 13: return statsMode(nums);
            case 14: return largeSmall(nums, k(a, 3, ec), true);
            case 15: return largeSmall(nums, k(a, 3, ec), false);
            // 20・21 は function_num として無効 (3 引数形でも #VALUE! なので上の範囲チェックで止まる)。
            default: return ErrorEval.NUM_ERROR;
        }
    }

    /** 関数番号 1〜11 (AVERAGE・COUNT・COUNTA・MAX・MIN・PRODUCT・STDEV・STDEVP・SUM・VAR・VARP) の集計 (nums は昇順)。それ以外は null。 */
    static ValueEval summary(int fn, List<Double> nums, int[] cnt) {
        switch (fn) {
            case 1: return statsMean(nums);
            case 2: return new NumberEval(cnt[0]);
            case 3: return new NumberEval(cnt[1]);
            case 4: return statsMax(nums);
            case 5: return statsMin(nums);
            case 6: return statsProduct(nums);
            case 7: return statsStdev(nums, false);
            case 8: return statsStdev(nums, true);
            case 9: return new NumberEval(statsSum(nums));
            case 10: return statsVar(nums, false);
            case 11: return statsVar(nums, true);
            default: return null;
        }
    }

    /**
     * SUBTOTAL(関数番号, 参照, ...)。1〜11 は手で隠した行も含め、101〜111 は隠した行を除く。
     * 入れ子の SUBTOTAL・AGGREGATE のセルは数えない。エラー値はそのまま伝える (COUNT・COUNTA は除く)。
     */
    static ValueEval subtotal(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length < 2)
            return ErrorEval.VALUE_INVALID;
        int fn = (int) Math.floor(num(a[0], ec));
        int base = fn > 100 ? fn - 100 : fn;
        if (base < 1 || base > 11)
            return ErrorEval.VALUE_INVALID;
        for (int i = 1; i < a.length; i++)
            if (!(a[i] instanceof AreaEval || a[i] instanceof RefEval || a[i] instanceof org.apache.poi.ss.formula.eval.RefListEval))
                return ErrorEval.VALUE_INVALID; // 参照だけ受け付ける
        Opts opts = new Opts(fn > 100, false, true, base != 2 && base != 3);
        List<Double> nums = new ArrayList<>();
        int[] cnt = new int[2];
        for (int i = 1; i < a.length; i++)
            aggCollect(a[i], true, opts, nums, cnt);
        java.util.Collections.sort(nums);
        return summary(base, nums, cnt);
    }

    /**
     * AGGREGATE の参照引数の検査 (実測):
     *  ・2 番目のデータが数値などの「定数」だと #VALUE! (AGGREGATE(9,0,1) = #VALUE!)
     *  ・k を取らない関数で 4 番目以降に「定数」を渡すと #VALUE! (AGGREGATE(9,0,A207:C207,1)・AGGREGATE(12,0,r,2))
     *  ・4 番目以降に「参照」を渡すとそのまま足し込まれる (AGGREGATE(9,0,A201:H201,A211:G211) = 47+34 = 81)
     */
    static boolean aggRefsOk(ValueEval[] a, boolean kFn) {
        for (int i = 2; i < a.length; i++) {
            if (kFn && i >= 3)
                return true; // 14..19 の 4 番目以降は k・q。ここでは数えない
            ValueEval v = a[i];
            // インライン配列定数 {2;3;7} は POI では LazyAreaEval ではなく **CacheAreaEval** として届く (実行時に型を見て実測)。
            // Excel は配列定数を AGGREGATE のデータに渡すと #VALUE! (AGGREGATE(9,0,{2;3;7}) = #VALUE!)なので、ここで弾く。
            if (v instanceof org.apache.poi.ss.formula.CacheAreaEval)
                return false;
            if (v instanceof AreaEval || v instanceof RefEval)
                continue;
            if (v instanceof org.apache.poi.ss.formula.eval.RefListEval rl) {
                // 参照のまとめ (union) は有効 (実測: AGGREGATE(9,0,A201:H201,A211:G211) = 81)。
                // 配列定数 {2;3;7} も POI では同じ RefListEval に載るので、中身が参照かどうかで判別する
                // (実測: AGGREGATE(9,0,{2;3;7}) = #VALUE!)。
                for (Object x : rl.getList())
                    if (!(x instanceof RefEval) && !(x instanceof AreaEval))
                        return false;
                continue;
            }
            return false;
        }
        return true;
    }

    /** AGGREGATE の第 4 引数 (k・q)。Excel は 1 以上・整数の範囲外を #NUM!、無いものは #VALUE! にする。 */
    static double k(ValueEval[] a, int at, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length <= at)
            throw new EvaluationException(ErrorEval.VALUE_INVALID);
        return num(a[at], ec);
    }

    /** PERCENTILE.INC / PERCENTILE.EXC の位置から値を出す (nums は昇順に済んでいる)。 */
    static ValueEval percentileFrom(List<Double> nums, double k, boolean exc) {
        if (nums.isEmpty() || Double.isNaN(k))
            return ErrorEval.NUM_ERROR;
        int n = nums.size();
        double p; // 0 基準の位置
        if (exc) {
            p = k * (n + 1) - 1;
            if (k <= 0 || k >= 1 || p < 0 || p > n - 1)
                return ErrorEval.NUM_ERROR;
        } else {
            if (k < 0 || k > 1)
                return ErrorEval.NUM_ERROR;
            p = k * (n - 1);
        }
        int lo = (int) Math.floor(p), hi = (int) Math.ceil(p);
        if (hi > n - 1)
            hi = n - 1;
        return new NumberEval(nums.get(lo) + (p - lo) * (nums.get(hi) - nums.get(lo)));
    }

    static ValueEval largeSmall(List<Double> nums, double kd, boolean largest) throws EvaluationException {
        if (Double.isNaN(kd))
            return ErrorEval.VALUE_INVALID;
        // 実測で決めた k の扱い (配列 {9,3,2,4,7,7,2} = 昇順 2,2,3,4,7,7,9・個数 7・隠し列 G を含む):
        //   k < 1 と 個数超え (#NUM!) を先に弾く: LARGE(...,0.999) = #NUM!、LARGE(...,8) = #NUM!。
        //   LARGE: 添字 = floor(n - k) (LARGE(...,1.9) = 7、LARGE(...,2) = 7、LARGE(...,1) = 9)
        //   SMALL: 添字 = floor(k) - 1 (SMALL(...,2.5) = 2、SMALL(...,3.5) = 3 ← 四捨五入ではなく切り捨て)
        //   同じ k で LARGE が 2 番目・SMALL が切り捨てになるのは、Excel が逆順の添字を切り捨てているために見える。
        if (kd < 1)
            return ErrorEval.NUM_ERROR;
        int n = nums.size();
        int idx = largest ? (int) Math.floor(n - kd) : (int) Math.floor(kd) - 1;
        if (idx < 0 || idx >= n)
            return ErrorEval.NUM_ERROR;
        return new NumberEval(nums.get(idx));
    }

    // ---- AGGREGATE の集計本体 (nums は昇順に済んでいる)。空のときの #DIV/0!・#NUM! は実測で確定する ----

    /** MEDIAN (function_num 12)。実測: AGGREGATE(12,0,A201:H201) = 4.5、AGGREGATE(12,0,A211:G211) = 4。数値無しは #NUM! (Excel の MEDIAN と同じ向き・未測定)。 */
    static ValueEval statsMedian(List<Double> v) {
        int n = v.size();
        if (n == 0)
            return ErrorEval.NUM_ERROR;
        return new NumberEval(n % 2 == 1 ? v.get(n / 2) : (v.get(n / 2 - 1) + v.get(n / 2)) / 2.0);
    }

    /**
     * MODE.SNGL (function_num 13)。意味を実測で確定した:
     *  AGGREGATE(13,0,A201:H201) = **#N/A** (1,2,3,4,5,7,10,15 はすべて 1 回ずつ = 最頻値なし)
     *  AGGREGATE(13,0,A211:G211) = **2** ({9,3,2,4,7,7,2} は 2 と 7 の同率最頻。昇順で最初に現れた方 = 小さい方を返す)
     */
    static ValueEval statsMode(List<Double> v) {
        int n = v.size();
        if (n == 0)
            return ErrorEval.NA;
        double best = v.get(0);
        int bestCnt = 1, cnt = 1;
        for (int i = 1; i < n; i++) {
            // Double 参照の == は箱入れ別の比較になるので必ず二重化する (実測バグで #N/A を返していた)
            cnt = v.get(i).doubleValue() == v.get(i - 1).doubleValue() ? cnt + 1 : 1;
            if (cnt > bestCnt) {
                bestCnt = cnt;
                best = v.get(i);
            }
        }
        return bestCnt == 1 ? ErrorEval.NA : new NumberEval(best);
    }

    static ValueEval statsMean(List<Double> v) {
        if (v.isEmpty())
            return ErrorEval.DIV_ZERO;
        double s = 0;
        for (double x : v)
            s += x;
        return new NumberEval(s / v.size());
    }

    static ValueEval statsMax(List<Double> v) {
        if (v.isEmpty())
            return new NumberEval(0); // 数値がなければ 0 (Excel)
        return new NumberEval(v.get(v.size() - 1));
    }

    static ValueEval statsMin(List<Double> v) {
        if (v.isEmpty())
            return new NumberEval(0);
        return new NumberEval(v.get(0));
    }

    static ValueEval statsProduct(List<Double> v) {
        if (v.isEmpty())
            return new NumberEval(0); // 数値がなければ 0 (Excel)
        double p = 1;
        for (double x : v)
            p *= x;
        return new NumberEval(p);
    }

    static double statsSum(List<Double> v) {
        double s = 0;
        for (double x : v)
            s += x;
        return s;
    }

    static double statsSsq(List<Double> v) {
        double m = statsSum(v) / v.size(), ss = 0;
        for (double x : v)
            ss += (x - m) * (x - m);
        return ss;
    }

    static ValueEval statsStdev(List<Double> v, boolean pop) {
        int n = v.size();
        if (pop ? n < 1 : n < 2)
            return ErrorEval.DIV_ZERO;
        return new NumberEval(Math.sqrt(statsSsq(v) / (pop ? n : n - 1)));
    }

    static ValueEval statsVar(List<Double> v, boolean pop) {
        int n = v.size();
        if (pop ? n < 1 : n < 2)
            return ErrorEval.DIV_ZERO;
        return new NumberEval(statsSsq(v) / (pop ? n : n - 1));
    }

    /** AGGREGATE 用の集め方。エラーは無視オプションのときだけ飛ばし、隠し行・隠し列を飛ばせる。 */
    /** 集め方の指定 (options から決まる)。 */
    record Opts(boolean ignoreHidden, boolean ignoreErrors, boolean ignoreNested, boolean propagateErrors) {
    }

    /**
     * データの集め方。範囲の中の数値だけを nums に入れ、cnt[0] = 数値の個数 (COUNT)、cnt[1] = 空でない個数 (COUNTA)。
     * 無視しないエラー値は、集計がそれを使う関数 (COUNT・COUNTA 以外) ならそのまま伝える。
     */
    static void aggCollect(ValueEval v, boolean fromRef, Opts o, List<Double> out, int[] cnt) throws EvaluationException {
        if (v instanceof ErrorEval e) {
            if (o.ignoreErrors)
                return;
            if (o.propagateErrors)
                throw new EvaluationException(e);
            cnt[1]++; // COUNTA はエラー値も数える
            return;
        }
        if (v instanceof AreaEval ar) {
            int sh = ar.getFirstSheetIndex();
            for (int r = ar.getFirstRow(); r <= ar.getLastRow(); r++) {
                if (o.ignoreHidden && rowHidden(sh, r))
                    continue;
                for (int c = ar.getFirstColumn(); c <= ar.getLastColumn(); c++) {
                    if (o.ignoreNested && isNestedAggregate(sh, r, c))
                        continue;
                    try {
                        aggCollect(ar.getAbsoluteValue(r, c), true, o, out, cnt);
                    } catch (EvaluationException e) {
                        throw e;
                    } catch (RuntimeException e) {
                        // 範囲の中のエラーセル (=NA() など) を読むと POI が例外を投げることがある。エラー値として扱う
                        aggCollect(ErrorEval.NA, true, o, out, cnt);
                    }
                }
            }
            return;
        }
        if (v instanceof RefEval re) {
            int sh = re.getFirstSheetIndex();
            if (o.ignoreHidden && rowHidden(sh, re.getRow()))
                return;
            if (o.ignoreNested && isNestedAggregate(sh, re.getRow(), re.getColumn()))
                return;
            aggCollect(re.getInnerValueEval(sh), true, o, out, cnt);
            return;
        }
        if (v instanceof org.apache.poi.ss.formula.eval.RefListEval rl) {
            for (Object x : rl.getList())
                aggCollect((ValueEval) x, true, o, out, cnt);
            return;
        }
        if (v == BlankEval.instance || v == org.apache.poi.ss.formula.eval.MissingArgEval.instance)
            return;
        cnt[1]++;
        if (v instanceof NumberEval n) {
            out.add(n.getNumberValue());
            cnt[0]++;
        } else if (fromRef) {
            // 参照の中の文字列・論理値は集めない (COUNTA では数える)
        } else if (v instanceof BoolEval b) {
            out.add(b.getBooleanValue() ? 1d : 0d);
            cnt[0]++;
        } else {
            out.add(OperandResolver.coerceValueToDouble(v));
            cnt[0]++;
        }
    }

    /** セルの数式が SUBTOTAL か AGGREGATE か (options 0〜3 は入れ子の集計を無視する)。 */
    static boolean isNestedAggregate(int sheetIdx, int row, int col) {
        try {
            XSSFWorkbook w = FormulaContext.workbook();
            int si = sheetIndexAt(sheetIdx);
            if (w == null || si < 0 || si >= w.getNumberOfSheets())
                return false;
            org.apache.poi.ss.usermodel.Row r = w.getSheetAt(si).getRow(row);
            org.apache.poi.ss.usermodel.Cell c = r == null ? null : r.getCell(col);
            if (c == null || c.getCellType() != org.apache.poi.ss.usermodel.CellType.FORMULA)
                return false;
            String f = c.getCellFormula().toUpperCase(Locale.ROOT);
            return f.contains("SUBTOTAL(") || f.contains("AGGREGATE(");
        } catch (RuntimeException e) {
            return false;
        }
    }

    static int sheetIndexAt(int idx) {
        return idx < 0 ? FormulaContext.currentSheet() : idx;
    }

    static boolean rowHidden(int sheetIdx, int row) {
        try {
            XSSFWorkbook w = FormulaContext.workbook();
            if (w == null)
                return false;
            int si = sheetIndexAt(sheetIdx);
            if (si < 0 || si >= w.getNumberOfSheets())
                return false;
            org.apache.poi.ss.usermodel.Row r = w.getSheetAt(si).getRow(row);
            return r != null && r.getZeroHeight();
        } catch (RuntimeException e) {
            return false;
        }
    }

}
