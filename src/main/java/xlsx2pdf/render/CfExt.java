package xlsx2pdf.render;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;
import java.util.function.Function;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtension;
import org.w3c.dom.Element;

/**
 * 条件付き書式の Excel 2010 以降の拡張 (シートの extLst の x14:conditionalFormatting) と、アイコンセット・データバーの
 * しきい値の計算の共通部分。POI は拡張を読まないので、シートの XML を読み直す。
 */
final class CfExt {
    private CfExt() {}

    /** x14 の条件付き書式の拡張の URI。 */
    static final String X14_CF = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

    /**
     * 拡張の規則 1 つ。
     *
     * @param rule   x14:cfRule 要素
     * @param ranges 適用範囲
     */
    record X14Rule(Element rule, List<CellRangeAddress> ranges) {}

    /** シートの拡張の規則 (なければ空)。 */
    static List<X14Rule> x14Rules(XSSFSheet sheet) {
        List<X14Rule> out = new ArrayList<>();
        var ws = sheet.getCTWorksheet();
        boolean x14 = false;
        if (ws.isSetExtLst()) for (CTExtension e : ws.getExtLst().getExtArray()) x14 |= X14_CF.equalsIgnoreCase(e.getUri());
        if (!x14) return out;
        Element root;
        try (java.io.InputStream in = sheet.getPackagePart().getInputStream()) {
            root = Xml.parse(in);
        } catch (Exception e) {
            Warnings.add("条件付き書式の拡張を読めませんでした", e);
            return out;
        }
        for (Element ext : Xml.children(Xml.child(root, "extLst"), "ext")) {
            if (!X14_CF.equalsIgnoreCase(ext.getAttribute("uri"))) continue;
            for (Element cfs : Xml.children(ext, "conditionalFormattings")) {
                for (Element cf : Xml.children(cfs, "conditionalFormatting")) {
                    List<CellRangeAddress> ranges = ranges(Xml.text(Xml.child(cf, "sqref")));
                    for (Element r : Xml.children(cf, "cfRule")) out.add(new X14Rule(r, ranges));
                }
            }
        }
        return out;
    }

    /** 通常の規則が拡張の規則を指す ID (extLst の x14:id)。なければ空。 */
    static List<String> x14Ids(CTCfRule r) {
        List<String> out = new ArrayList<>();
        if (!r.isSetExtLst()) return out;
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("<(?:\\w+:)?id>([^<]+)</").matcher(r.getExtLst().xmlText());
        while (m.find()) out.add(m.group(1).trim());
        return out;
    }

    /** "A1:B3 D5" の形の範囲の一覧。 */
    static List<CellRangeAddress> ranges(String sqref) {
        List<CellRangeAddress> out = new ArrayList<>();
        if (sqref == null) return out;
        for (String p : sqref.trim().split("\\s+")) {
            if (p.isEmpty()) continue;
            try {
                out.add(CellRangeAddress.valueOf(p));
            } catch (RuntimeException e) {
                Warnings.add("条件付き書式の範囲 " + p + " を読めませんでした", e);
            }
        }
        return out;
    }

    /** sqref (POI の型は List か文字列) を文字列にする。 */
    static String sqref(Object o) {
        return o instanceof List<?> l ? String.join(" ", l.stream().map(String::valueOf).toList()) : String.valueOf(o);
    }

    /** 範囲のどれかに入るか。 */
    static boolean in(List<CellRangeAddress> ranges, int r, int c) {
        for (CellRangeAddress a : ranges) if (a.isInRange(r, c)) return true;
        return false;
    }

    /** 適用範囲の数値 (昇順)。範囲が列全体 (A:A) でも、ある行・セルだけを調べる。 */
    static double[] values(XSSFSheet sheet, List<CellRangeAddress> ranges, BiFunction<Integer, Integer, Double> value) {
        List<Double> vals = new ArrayList<>();
        for (org.apache.poi.ss.usermodel.Row row : sheet) {
            int r = row.getRowNum();
            for (org.apache.poi.ss.usermodel.Cell cell : row) {
                int c = cell.getColumnIndex();
                if (in(ranges, r, c)) {
                    Double v = value.apply(r, c);
                    if (v != null) vals.add(v);
                }
            }
        }
        double[] arr = new double[vals.size()];
        for (int i = 0; i < arr.length; i++) arr[i] = vals.get(i);
        Arrays.sort(arr);
        return arr;
    }

    /**
     * しきい値 (cfvo) の値。
     *
     * @param type   num, percent, percentile, min, max, autoMin, autoMax, formula
     * @param val    値の文字列 (数値か数式)
     * @param sorted 適用範囲の数値 (昇順)
     * @return 値。決められなければ NaN
     */
    static double threshold(String type, String val, double[] sorted, Function<String, Double> formula) {
        double min = sorted.length == 0 ? 0 : sorted[0], max = sorted.length == 0 ? 0 : sorted[sorted.length - 1];
        Double x = number(val, formula);
        return switch (type) {
            case "percent" -> x == null ? Double.NaN : min + (max - min) * x / 100;
            case "percentile" -> x == null ? Double.NaN : percentile(sorted, x / 100);
            case "min" -> min;
            case "max" -> max;
            case "autoMin" -> Math.min(0, min);
            case "autoMax" -> Math.max(0, max);
            default -> x == null ? Double.NaN : x; // num, formula
        };
    }

    /** しきい値の文字列を数値にする (数式なら評価する)。 */
    static Double number(String s, Function<String, Double> formula) {
        if (s == null || s.isBlank()) return null;
        String t = s.trim();
        if (t.startsWith("=")) t = t.substring(1);
        try {
            return Double.parseDouble(t);
        } catch (NumberFormatException e) {
            return formula.apply(t);
        }
    }

    /** PERCENTILE.INC と同じ補間の百分位数。 */
    static double percentile(double[] sorted, double p) {
        if (sorted.length == 0) return Double.NaN;
        double k = Math.max(0, Math.min(1, p)) * (sorted.length - 1);
        int i = (int) Math.floor(k);
        double f = k - i;
        return i + 1 < sorted.length ? sorted[i] + f * (sorted[i + 1] - sorted[i]) : sorted[i];
    }

    /** ID で引く表を作る補助。 */
    static <T> Map<String, T> map() { return new HashMap<>(); }
}
