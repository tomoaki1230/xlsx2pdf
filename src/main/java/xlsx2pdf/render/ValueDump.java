package xlsx2pdf.render;

import java.util.ArrayList;
import java.util.List;

/**
 * 配置の数値の書き出し (開発用)。システムプロパティ xlsx2pdf.dumpValues があるとき、変換ごとに
 * 「出力の PDF のパス + .values.tsv」へ、Excel の scripts/dump-excel.ps1 と同じ形 (シート<TAB>種類<TAB>値...) で書く。
 * Excel の値と突き合わせて、ずれた箇所を自動で探すために使う (設計書 12.2)。
 */
public final class ValueDump {
    /** 書き出すか。 */
    public static final boolean ON = System.getProperty("xlsx2pdf.dumpValues") != null;
    private static final ThreadLocal<List<String>> LINES = new ThreadLocal<>();
    private static final ThreadLocal<String> SHEET = new ThreadLocal<>();

    private ValueDump() {
    }

    /** 変換の始まり。 */
    public static void begin() {
        if (ON) LINES.set(new ArrayList<>());
    }

    /** これから書く値のシート名。 */
    public static void sheet(String name) {
        if (ON) SHEET.set(name);
    }

    /** 1 行を足す (数値は小数 2 桁)。 */
    public static void add(String kind, Object... vals) {
        List<String> l = ON ? LINES.get() : null;
        if (l == null) return;
        StringBuilder sb = new StringBuilder(String.valueOf(SHEET.get())).append('\t').append(kind);
        for (Object v : vals) sb.append('\t').append(v instanceof Double d ? fmt(d) : String.valueOf(v));
        l.add(sb.toString());
    }

    /** 小数 2 桁の文字列 (整数なら小数なし)。 */
    public static String fmt(double d) {
        double r = Math.round(d * 100) / 100.0;
        return r == Math.rint(r) ? String.valueOf((long) r) : String.valueOf(r);
    }

    /** 変換の終わり: 集めた行を返して消す。 */
    public static List<String> take() {
        List<String> l = LINES.get();
        LINES.remove();
        SHEET.remove();
        return l == null ? List.of() : l;
    }
}
