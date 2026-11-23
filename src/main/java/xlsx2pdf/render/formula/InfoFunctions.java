package xlsx2pdf.render.formula;

import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.eval.BoolEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static xlsx2pdf.render.formula.FnSupport.*;

/** 情報・論理の関数 (ISFORMULA・N・TYPE・XOR)。 */
final class InfoFunctions {
    private InfoFunctions() {
    }

    // ---- ISFORMULA・N・TYPE ----

    /**
     * ISFORMULA(参照)。実 Excel では範囲を渡すと「1 つでも式を含むセルがあれば TRUE」
     * (ISFORMULA(A350:D350) = TRUE と実測)。1 セルの参照でも同じ規則で数える。
     */
    static ValueEval isFormula(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        XSSFWorkbook wb = FormulaContext.workbook();
        if (wb == null) return BoolEval.FALSE;
        if (a[0] instanceof AreaEval ar) {
            int ash = ar.getFirstSheetIndex();
            if (ash < 0 || ash >= wb.getNumberOfSheets()) return BoolEval.FALSE;
            // 範囲は左上のセルで調べる (Excel の動的配列では範囲全体の配列が返り、1 つのセルには先頭が入る)
            return BoolEval.valueOf(hasFormula(wb.getSheetAt(ash), ar.getFirstRow(), ar.getFirstColumn()));
        }
        if (a[0] instanceof RefEval re) {
            int rsh = re.getFirstSheetIndex();
            if (rsh < 0 || rsh >= wb.getNumberOfSheets()) return BoolEval.FALSE;
            return BoolEval.valueOf(hasFormula(wb.getSheetAt(rsh), re.getRow(), re.getColumn()));
        }
        return ErrorEval.VALUE_INVALID;
    }

    static boolean hasFormula(XSSFSheet s, int row, int col) {
        if (row < 0 || col < 0) return false;
        org.apache.poi.ss.usermodel.Row r = s.getRow(row);
        if (r == null) return false;
        org.apache.poi.ss.usermodel.Cell c = r.getCell(col);
        return c != null && c.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA;
    }

    /**
     * N(値)。数値はその値、日付はシリアル値、論理値は 1/0、文字列・空は 0 (Excel)。
     * 範囲を渡すと先頭のセルで評価する。
     */
    static ValueEval n(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        ValueEval v = a[0] instanceof AreaEval ar ? ar.getRelativeValue(0, 0) : scalar(a[0], ec);
        if (v instanceof NumberEval) return v;
        if (v instanceof BoolEval b) return new NumberEval(b.getBooleanValue() ? 1 : 0);
        if (v instanceof ErrorEval e) throw new EvaluationException(e);
        if (v == BlankEval.instance) return new NumberEval(0);
        if (v instanceof StringEval) return new NumberEval(0);
        return ErrorEval.VALUE_INVALID;
    }

    /**
     * TYPE(値): 1 数値, 2 文字列, 4 論理値, 16 エラー, 64 配列。
     * 2 つ以上のセルの範囲と配列定数は 64、1 つのセルの範囲はそのセルの種類。
     */
    static ValueEval type(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        ValueEval v = a[0];
        if (v instanceof AreaEval ar) {
            if (ar.getHeight() * ar.getWidth() > 1) return new NumberEval(64);
            v = ar.getRelativeValue(0, 0);
        } else if (v instanceof org.apache.poi.ss.formula.CacheAreaEval || v instanceof org.apache.poi.ss.formula.eval.RefListEval) {
            return new NumberEval(64);
        } else if (v instanceof RefEval re) v = re.getInnerValueEval(re.getFirstSheetIndex());
        if (v instanceof NumberEval || v == BlankEval.instance) return new NumberEval(1);
        if (v instanceof StringEval) return new NumberEval(2);
        if (v instanceof BoolEval) return new NumberEval(4);
        if (v instanceof ErrorEval) return new NumberEval(16);
        return new NumberEval(1);
    }

    /** XOR: 真の個数が奇数なら TRUE。論理値でない文字列は無視し、論理値が 1 つもなければ #VALUE!。 */
    static ValueEval xor(ValueEval[] args, OperationEvaluationContext ec) throws EvaluationException {
        int[] st = new int[2]; // [真の数, 論理として読めた数]
        for (ValueEval a : args) {
            if (a instanceof AreaEval ar) {
                for (int r = 0; r < ar.getHeight(); r++)
                    for (int c = 0; c < ar.getWidth(); c++) xorOne(ar.getRelativeValue(r, c), true, st);
            } else if (a instanceof RefEval re) {
                xorOne(re.getInnerValueEval(re.getFirstSheetIndex()), true, st);
            } else {
                xorOne(a, false, st);
            }
        }
        if (st[1] == 0) return ErrorEval.VALUE_INVALID;
        return BoolEval.valueOf(st[0] % 2 == 1);
    }

    static void xorOne(ValueEval v, boolean fromRef, int[] st) throws EvaluationException {
        if (v instanceof ErrorEval e) throw new EvaluationException(e);
        if (v instanceof BoolEval b) {
            st[1]++;
            if (b.getBooleanValue()) st[0]++;
        } else if (v instanceof NumberEval n) {
            st[1]++;
            if (n.getNumberValue() != 0) st[0]++;
        } else if (v instanceof StringEval s && !fromRef) {
            String t = s.getStringValue();
            if (t.equalsIgnoreCase("TRUE")) {
                st[1]++;
                st[0]++;
            } else if (t.equalsIgnoreCase("FALSE")) {
                st[1]++;
            }
            // 論理値でない文字列は、Excel と同じく無視する
        }
    }
}
