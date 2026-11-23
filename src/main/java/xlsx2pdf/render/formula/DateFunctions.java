package xlsx2pdf.render.formula;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.temporal.IsoFields;
import java.util.Locale;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import static xlsx2pdf.render.formula.FnSupport.*;

/** 日付の関数 (ISOWEEKNUM・DATEDIF) と、Excel の暦 (1900 年 2 月 29 日あり) の変換。 */
final class DateFunctions {
    private DateFunctions() {
    }


    /** Excel の暦 (1900 年 2 月 29 日を含む) の年月日。 */
    record Ymd(int y, int m, int d) {
    }

    /** 日付 (シリアル値の整数部) の暦の日。 */
    static Ymd ymd(long n) throws EvaluationException {
        if (n < 0) throw new EvaluationException(ErrorEval.NUM_ERROR);
        if (FormulaContext.is1904()) {
            LocalDate l = LocalDate.of(1904, 1, 1).plusDays(n);
            return new Ymd(l.getYear(), l.getMonthValue(), l.getDayOfMonth());
        }
        if (n == 0) return new Ymd(1900, 1, 0);
        if (n == 60) return new Ymd(1900, 2, 29);
        LocalDate l = n < 60 ? LocalDate.of(1899, 12, 31).plusDays(n) : LocalDate.of(1899, 12, 30).plusDays(n);
        return new Ymd(l.getYear(), l.getMonthValue(), l.getDayOfMonth());
    }

    static long serial(Ymd t) {
        if (FormulaContext.is1904()) return java.time.temporal.ChronoUnit.DAYS.between(LocalDate.of(1904, 1, 1), LocalDate.of(t.y, t.m, t.d));
        if (t.y == 1900 && t.m == 2 && t.d == 29) return 60;
        long n = java.time.temporal.ChronoUnit.DAYS.between(LocalDate.of(1899, 12, 30), LocalDate.of(t.y, t.m, t.d));
        return n < 61 ? n - 1 : n;
    }

    /** Excel の暦での月の日数 (1900 年は 2 月が 29 日)。 */
    static int monthLen(int y, int m) {
        if (y == 1900 && m == 2 && !FormulaContext.is1904()) return 29;
        return YearMonth.of(y, m).lengthOfMonth();
    }

    /** 日付の引数 (数値・日付に読める文字列)。 */
    static double dateArg(ValueEval v, OperationEvaluationContext ec) throws EvaluationException {
        return dateArg(OperandResolver.getSingleValue(v, FnSupport.rowOf(ec), FnSupport.colOf(ec)));
    }

    static double dateArg(ValueEval v) throws EvaluationException {
        if (v instanceof StringEval se) {
            String t = se.getStringValue().trim();
            java.util.regex.Matcher m = java.util.regex.Pattern.compile("(\\d{4})[/\\-.年](\\d{1,2})[/\\-.月](\\d{1,2})日?").matcher(t);
            if (!m.matches()) m = java.util.regex.Pattern.compile("(\\d{1,2})/(\\d{1,2})/(\\d{4})").matcher(t);
            if (m.matches()) {
                try {
                    boolean ymdFirst = m.group(1).length() == 4;
                    int y = Integer.parseInt(ymdFirst ? m.group(1) : m.group(3));
                    int mo = Integer.parseInt(ymdFirst ? m.group(2) : m.group(1));
                    int d = Integer.parseInt(ymdFirst ? m.group(3) : m.group(2));
                    return serial(new Ymd(y, mo, d));
                } catch (RuntimeException e) {
                    throw new EvaluationException(ErrorEval.VALUE_INVALID);
                }
            }
        }
        return OperandResolver.coerceValueToDouble(v);
    }

    static ValueEval isoWeekNum(ValueEval[] a, OperationEvaluationContext ec) throws EvaluationException {
        if (a.length != 1) return ErrorEval.VALUE_INVALID;
        double n = dateArg(a[0], ec);
        if (n < 0) return ErrorEval.NUM_ERROR;
        // Excel は 1900 年 1 月 1 日 (シリアル値 1) を日曜日とみなす。曜日は暦どおりに 1899/12/30 を 0 として数える
        LocalDate base = FormulaContext.is1904() ? LocalDate.of(1904, 1, 1) : LocalDate.of(1899, 12, 30);
        return new NumberEval(base.plusDays((long) Math.floor(n)).get(IsoFields.WEEK_OF_WEEK_BASED_YEAR));
    }

    static ValueEval datedif(ValueEval[] a, int row, int col) throws EvaluationException {
        if (a.length != 3) return ErrorEval.VALUE_INVALID;
        double s = dateArg(OperandResolver.getSingleValue(a[0], row, col));
        double e = dateArg(OperandResolver.getSingleValue(a[1], row, col));
        String unit = OperandResolver.coerceValueToString(OperandResolver.getSingleValue(a[2], row, col)).toUpperCase(Locale.ROOT);
        long n1 = (long) Math.floor(s), n2 = (long) Math.floor(e);
        if (n1 < 0 || n2 < 0) return ErrorEval.NUM_ERROR;
        if (n1 > n2) return ErrorEval.NUM_ERROR;
        Ymd d1 = ymd(n1), d2 = ymd(n2);
        int months = (d2.y - d1.y) * 12 + d2.m - d1.m;
        if (d2.d < d1.d) months--;
        switch (unit) {
            case "Y":
                return new NumberEval(months / 12);
            case "M":
                return new NumberEval(months);
            case "D":
                return new NumberEval(n2 - n1);
            case "YM":
                return new NumberEval(months % 12);
            case "MD": {
                int d = d2.d - d1.d;
                if (d < 0) {
                    int pm = d2.m == 1 ? 12 : d2.m - 1, py = d2.m == 1 ? d2.y - 1 : d2.y;
                    d += monthLen(py, pm);
                }
                return new NumberEval(d);
            }
            case "YD": {
                // Excel の結果に合わせた: 開始日に「月の端数 (YM) か月」を足した日 (31 日などは翌月へあふれる) までの日数に、MD の日数を足す。
                // 足す先は開始日の年の暦のまま (開始日の年がうるう年なら 2 月 29 日を数える)
                int ym = months % 12;
                int md = d2.d - d1.d;
                if (md < 0) {
                    int pm = d2.m == 1 ? 12 : d2.m - 1, py = d2.m == 1 ? d2.y - 1 : d2.y;
                    md += monthLen(py, pm);
                }
                LocalDate st = real(d1);
                LocalDate t = LocalDate.of(st.getYear(), 1, 1).plusMonths(st.getMonthValue() - 1 + ym).plusDays(st.getDayOfMonth() - 1);
                long days = java.time.temporal.ChronoUnit.DAYS.between(st, t) + md;
                // 1900 年は Excel では 2 月 29 日がある
                if (!FormulaContext.is1904() && d1.y == 1900 && st.isBefore(LocalDate.of(1900, 3, 1)) && !t.isBefore(LocalDate.of(1900, 3, 1))) days++;
                return new NumberEval(days);
            }
            default:
                return ErrorEval.NUM_ERROR;
        }
    }

    static LocalDate real(Ymd t) {
        if (t.d == 0) return LocalDate.of(t.y, t.m, 1).minusDays(1);
        if (t.y == 1900 && t.m == 2 && t.d == 29) return LocalDate.of(1900, 3, 1);
        return LocalDate.of(t.y, t.m, t.d);
    }
}
