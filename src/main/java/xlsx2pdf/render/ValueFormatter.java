package xlsx2pdf.render;

import java.awt.Color;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

/** セルについて Excel が表示する文字列を作る (表示形式の適用は {@link ExcelFormat})。 */
public final class ValueFormatter {
    /** 表示する文字列と、その表示に関する情報。 */
    public static final class Result {
        /** 表示する文字列。 */
        public String text = "";
        /** 表示形式の [色] の指定 (null ならフォントの色)。 */
        public Color color;
        /** 数値 (右揃えで、隣のセルにはみ出さない)。 */
        public boolean number;
        /** 真偽値・エラー値 (標準の配置では中央揃え)。 */
        public boolean bool;
        /** 真偽値 (TRUE/FALSE。収まらなければ数値と同じく # で埋める)。 */
        public boolean logical;
        /** 標準 (General) の表示形式 (幅に合わせて桁数を減らせる)。 */
        public boolean general;
        /** 数値の値。 */
        public double value;
        /** 繰り返し文字 ("*x" の x。なければ null)。 */
        public String fillChar;
        /** 繰り返し位置の前・後の文字列。 */
        public String fillLeft, fillRight;
        /** 書式付き文字列 (文字列セルで書式の区切りがあるとき)。 */
        public XSSFRichTextString rich;
    }

    /** {@link ExcelFormat}: 表示形式の書式化。 */
    private final ExcelFormat excel;
    /** 日本語のブックか (組み込み表示形式の日本語版を使う)。 */
    private final boolean japanese;
    /** ブックの表示形式の定義。 */
    private final StylesTable styles;

    /** 日本語版 Excel の組み込み表示形式 (ファイルには書かれない)。 */
    private static final Map<Integer, String> JA_BUILTIN = new HashMap<>();

    static {
        // 通貨と桁区切りも日本語版では負号の形で、括弧の余白がない (英語版は "#,##0_);[Red](#,##0)")
        JA_BUILTIN.put(5, "\"¥\"#,##0;\"¥\"\\-#,##0");
        JA_BUILTIN.put(6, "\"¥\"#,##0;[Red]\"¥\"\\-#,##0");
        JA_BUILTIN.put(7, "\"¥\"#,##0.00;\"¥\"\\-#,##0.00");
        JA_BUILTIN.put(8, "\"¥\"#,##0.00;[Red]\"¥\"\\-#,##0.00");
        JA_BUILTIN.put(37, "#,##0;\\-#,##0");
        JA_BUILTIN.put(38, "#,##0;[Red]\\-#,##0");
        JA_BUILTIN.put(39, "#,##0.00;\\-#,##0.00");
        JA_BUILTIN.put(40, "#,##0.00;[Red]\\-#,##0.00");
        JA_BUILTIN.put(41, "_ * #,##0_ ;_ * \\-#,##0_ ;_ * \"-\"_ ;_ @_ ");
        JA_BUILTIN.put(42, "_ \"¥\"* #,##0_ ;_ \"¥\"* \\-#,##0_ ;_ \"¥\"* \"-\"_ ;_ @_ ");
        JA_BUILTIN.put(43, "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * \"-\"??_ ;_ @_ ");
        JA_BUILTIN.put(44, "_ \"¥\"* #,##0.00_ ;_ \"¥\"* \\-#,##0.00_ ;_ \"¥\"* \"-\"??_ ;_ @_ ");
        JA_BUILTIN.put(14, "yyyy/m/d");
        JA_BUILTIN.put(22, "yyyy/m/d h:mm");
        JA_BUILTIN.put(27, "[$-411]ge.m.d");
        JA_BUILTIN.put(28, "[$-411]ggge\"年\"m\"月\"d\"日\"");
        JA_BUILTIN.put(29, "[$-411]ggge\"年\"m\"月\"d\"日\"");
        JA_BUILTIN.put(30, "m/d/yy");
        JA_BUILTIN.put(31, "yyyy\"年\"m\"月\"d\"日\"");
        JA_BUILTIN.put(32, "h\"時\"mm\"分\"");
        JA_BUILTIN.put(33, "h\"時\"mm\"分\"ss\"秒\"");
        JA_BUILTIN.put(34, "yyyy\"年\"m\"月\"");
        JA_BUILTIN.put(35, "m\"月\"d\"日\"");
        JA_BUILTIN.put(36, "[$-411]ge.m.d");
        JA_BUILTIN.put(50, "[$-411]ge.m.d");
        JA_BUILTIN.put(51, "[$-411]ggge\"年\"m\"月\"d\"日\"");
        JA_BUILTIN.put(52, "yyyy\"年\"m\"月\"");
        JA_BUILTIN.put(53, "m\"月\"d\"日\"");
        JA_BUILTIN.put(54, "[$-411]ggge\"年\"m\"月\"d\"日\"");
        JA_BUILTIN.put(55, "yyyy\"年\"m\"月\"");
        JA_BUILTIN.put(56, "m\"月\"d\"日\"");
        JA_BUILTIN.put(57, "[$-411]ge.m.d");
        JA_BUILTIN.put(58, "[$-411]ggge\"年\"m\"月\"d\"日\"");
    }

    /** 表示形式の [色] の英語名。 */
    private static final Map<String, Color> NAMED = Map.of(
            "black", Color.BLACK, "blue", Color.BLUE, "cyan", Color.CYAN, "green", new Color(0, 255, 0),
            "magenta", Color.MAGENTA, "red", Color.RED, "white", Color.WHITE, "yellow", Color.YELLOW);

    /** 表示形式の [ColorN] のパレット (N = 1〜56)。 */
    private static final int[] PALETTE = {0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
            0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC,
            0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000,
            0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF,
            0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300,
            0x993300, 0x993366, 0x333399, 0x333333};

    /** 角括弧の指定 ([赤] など)。 */
    private static final Pattern BRACKET = Pattern.compile("\\[([^\\]]+)\\]");

    /**
     * セル値の書式化器を作る。
     *
     * @param japanese 日本語のブックか
     * @param date1904 1904 年基準の日付か
     * @param styles   ブックの表示形式の定義
     */
    public ValueFormatter(boolean japanese, boolean date1904, StylesTable styles) {
        this.styles = styles;
        this.excel = new ExcelFormat(date1904);
        this.japanese = japanese;
    }

    /** セルの書式の表示形式を返す。日本語のブックでは、ファイルにない組み込み表示形式を日本語版にする。 */
    public String formatString(XSSFCellStyle st) {
        int idx = st.getDataFormat();
        if (japanese && idx < 164 && JA_BUILTIN.containsKey(idx)) {
            // 組み込み表示形式はファイルに書かれない。日本語版 Excel は独自の形式を使う
            String stored = styles.getNumberFormats().get((short) idx);
            if (stored == null) return JA_BUILTIN.get(idx);
        }
        String f = st.getDataFormatString();
        return f == null ? "General" : f;
    }

    /**
     * セルの表示文字列を作る。
     *
     * @param showZeros ゼロ値を表示するか
     * @param errorsAs  エラー値の印刷方法 (displayed, blank, dash, NA)
     */
    public Result format(Cell cell, boolean showZeros, String errorsAs) {
        return format(cell, showZeros, errorsAs, null);
    }

    /** 条件付き書式の表示形式 (fmtOverride) を指定して、セルの値を文字列にする。 */
    public Result format(Cell cell, boolean showZeros, String errorsAs, String fmtOverride) {
        Result r = new Result();
        CellType t = cell.getCellType();
        if (t == CellType.FORMULA) t = cell.getCachedFormulaResultType();
        XSSFCellStyle st = (XSSFCellStyle) cell.getCellStyle();
        String fmt = fmtOverride != null ? fmtOverride : formatString(st);
        switch (t) {
            case STRING -> {
                XSSFRichTextString rs = (XSSFRichTextString) cell.getRichStringCellValue();
                String s = rs.getString();
                r.text = s == null ? "" : s;
                if (rs.numFormattingRuns() > 0) r.rich = rs;
                applyTextSection(r, fmt);
            }
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                // 表示形式「@」の数値は文字列と同じ扱い (標準の配置で左寄せ)。
                // @ を含む区画が 1 つだけの書式 (@"円"・"No."@ など) も同じく、Excel は付けた文字を出さず「標準」の形で左寄せにする
                boolean text = fmt != null && (fmt.trim().equals("@") || textOnly(fmt));
                r.number = !text;
                if (text) fmt = "@";
                r.value = v;
                if (v == 0 && !showZeros) { r.text = ""; return r; }
                formatNumber(r, v, fmt);
            }
            case BOOLEAN -> {
                r.bool = true;
                r.logical = true;
                r.text = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            }
            case ERROR -> {
                r.bool = true;
                String e;
                try {
                    e = FormulaError.forInt(cell.getErrorCellValue()).getString();
                } catch (RuntimeException ex) {
                    e = "#VALUE!";
                }
                r.text = switch (errorsAs == null ? "displayed" : errorsAs) {
                    case "blank" -> "";
                    case "dash" -> "--";
                    case "NA" -> "#N/A";
                    default -> e;
                };
            }
            default -> r.text = "";
        }
        return r;
    }

    /** 文字列セルには表示形式の第 4 セクション ("@" の位置に文字列) を適用する。 */
    private void applyTextSection(Result r, String fmt) {
        if (fmt == null || fmt.equals("General") || fmt.equals("@")) return;
        ExcelFormat.Out o = excel.formatText(r.text, fmt);
        if (o == null) return;
        r.color = o.color;
        int f = o.text.indexOf(ExcelFormat.FILL);
        if (f >= 0) {
            r.fillChar = String.valueOf(o.fill == null ? ' ' : o.fill);
            r.fillLeft = o.text.substring(0, f);
            r.fillRight = o.text.substring(f + 1);
            r.text = r.fillLeft + r.fillRight;
            r.rich = null;
        } else {
            if (!o.text.equals(r.text)) r.rich = null;
            r.text = o.text;
        }
    }

    /** 数値を Excel の表示形式で文字列にする (グラフのラベル用)。 */
    public String formatNumber(double v, String fmt) {
        Result r = new Result();
        formatNumber(r, v, fmt == null || fmt.isEmpty() ? "General" : fmt);
        return r.fillLeft != null ? (r.fillLeft + r.fillRight).trim() : r.text;
    }

    /** 数値を書式化して結果に設定する。 */
    private void formatNumber(Result r, double v, String fmt) {
        // 文字列の表示形式「@」の数値 (形式を後から付けたセル) は、Excel は標準の形で表示する
        if (fmt == null || fmt.isEmpty() || fmt.trim().equals("@") || textOnly(fmt) || ExcelFormat.isGeneral(fmt) && sections(fmt).size() == 1
                && fmt.trim().equalsIgnoreCase("General")) {
            r.general = true;
            r.text = general(v, 11);
            return;
        }
        ExcelFormat.Out o = excel.format(v, fmt);
        r.color = o.color;
        // 「標準」の扱いは、値に使う区画が先頭の General のときだけ (General;[Red](#,##0) の負の値は 2 番目の区画)
        int ns = sections(fmt).size();
        if (ExcelFormat.isGeneral(fmt) && (ns == 1 || v > 0 || v == 0 && ns < 3)) r.general = true;
        int f = o.text.indexOf(ExcelFormat.FILL);
        if (f >= 0) {
            r.fillChar = String.valueOf(o.fill == null ? ' ' : o.fill);
            r.fillLeft = o.text.substring(0, f);
            r.fillRight = o.text.substring(f + 1).replace(String.valueOf(ExcelFormat.FILL), "");
            r.text = r.fillLeft + r.fillRight;
        } else {
            r.text = o.text;
        }
    }

    /** 表示形式をセクション (; 区切り。引用符・角括弧・エスケープの中は区切らない) に分ける。 */
    /** 区画が 1 つで、引用符の外に @ がある書式か (「@」だけの書式を除く)。 */
    public static boolean textOnly(String fmt) {
        return fmt != null && !fmt.trim().equals("@") && sections(fmt).size() == 1 && ExcelFormat.containsUnquoted(fmt, '@');
    }

    static List<String> sections(String fmt) {
        List<String> out = new ArrayList<>();
        StringBuilder cur = new StringBuilder();
        boolean q = false;
        int br = 0;
        for (int i = 0; i < fmt.length(); i++) {
            char c = fmt.charAt(i);
            if (c == '"') q = !q;
            else if (!q && c == '\\' && i + 1 < fmt.length()) { cur.append(c).append(fmt.charAt(++i)); continue; }
            else if (!q && c == '[') br++;
            else if (!q && c == ']') br--;
            else if (!q && br == 0 && c == ';') { out.add(cur.toString()); cur.setLength(0); continue; }
            cur.append(c);
        }
        out.add(cur.toString());
        return out;
    }

    /** セクションの [色] の指定 (英語名、ColorN、日本語名) を返す (なければ null)。 */
    static Color sectionColor(String sec) {
        Matcher m = BRACKET.matcher(sec);
        while (m.find()) {
            String n = m.group(1).toLowerCase(Locale.ROOT);
            Color c = NAMED.get(n);
            if (c != null) return c;
            if (n.startsWith("color")) {
                try {
                    int i = Integer.parseInt(n.substring(5).trim());
                    if (i >= 1 && i <= PALETTE.length) return new Color(PALETTE[i - 1]);
                } catch (NumberFormatException ignored) {
                    // 色の指定ではない
                }
            }
            switch (n) {
                case "赤" -> { return Color.RED; }
                case "青" -> { return Color.BLUE; }
                case "黒" -> { return Color.BLACK; }
                case "緑" -> { return new Color(0, 255, 0); }
                case "白" -> { return Color.WHITE; }
                case "黄" -> { return Color.YELLOW; }
                case "紫" -> { return Color.MAGENTA; }
                case "水" -> { return Color.CYAN; }
                default -> { }
            }
        }
        return null;
    }

    /**
     * 和暦を返す。
     *
     * @return {元号, 略号 (R, H, S, T, M), 年}。明治より前は {"", "", 西暦年}
     */
    static String[] era(LocalDate d) {
        if (!d.isBefore(LocalDate.of(2019, 5, 1))) return new String[] {"令和", "R", String.valueOf(d.getYear() - 2018)};
        if (!d.isBefore(LocalDate.of(1989, 1, 8))) return new String[] {"平成", "H", String.valueOf(d.getYear() - 1988)};
        if (!d.isBefore(LocalDate.of(1926, 12, 25))) return new String[] {"昭和", "S", String.valueOf(d.getYear() - 1925)};
        if (!d.isBefore(LocalDate.of(1912, 7, 30))) return new String[] {"大正", "T", String.valueOf(d.getYear() - 1911)};
        if (!d.isBefore(LocalDate.of(1868, 1, 1))) return new String[] {"明治", "M", String.valueOf(d.getYear() - 1867)};
        return new String[] {"", "", String.valueOf(d.getYear())};
    }

    /**
     * Excel の「標準」の表示形式で数値を文字列にする。
     *
     * @param maxChars 最大の文字数 (収まらなければ指数表示にする)
     */
    public static String general(double v, int maxChars) {
        if (v == 0) return "0";
        if (Double.isNaN(v) || Double.isInfinite(v)) return "#NUM!";
        double a = Math.abs(v);
        String s;
        if (a >= 1e-9 && a < 1e11) {
            int intDigits = a >= 1 ? (int) Math.floor(Math.log10(a)) + 1 : 1;
            int avail = maxChars - intDigits - (v < 0 ? 1 : 0) - 1;
            int dec = Math.max(0, Math.min(avail, 10 - Math.max(intDigits, 0) + (a < 1 ? leadingZeros(a) : 0)));
            dec = Math.min(dec, 30);
            java.math.BigDecimal bd = new java.math.BigDecimal(v).setScale(dec, java.math.RoundingMode.HALF_UP);
            // Excel と同じく有効数字 15 桁まで
            bd = bd.round(new java.math.MathContext(15));
            s = bd.stripTrailingZeros().toPlainString();
            if (s.length() <= maxChars && !(a < 1 && s.replace("-", "").equals("0"))) return s;
        }
        // 指数表示
        for (int digits = Math.max(0, maxChars - 6 - (v < 0 ? 1 : 0)); digits >= 0; digits--) {
            String m = String.format(Locale.ROOT, "%." + digits + "E", v);
            int e = m.indexOf('E');
            String mant = m.substring(0, e);
            if (mant.contains(".")) mant = mant.replaceAll("0+$", "").replaceAll("\\.$", "");
            int exp = Integer.parseInt(m.substring(e + 1));
            s = mant + "E" + (exp < 0 ? "-" : "+") + String.format("%02d", Math.abs(exp));
            if (s.length() <= maxChars) return s;
        }
        return String.format(Locale.ROOT, "%.0E", v);
    }

    /**
     * 狭い列の「標準」: maxChars 文字 (どの文字も数字 1 文字分の幅で数える) に収まる形。
     * 小数で書くと、1 未満の数の最初の有効数字の位置が変わるもの (0.000988 → 0.001) と、有効数字が指数表示より少ないものは指数表示にする。
     * どちらも収まらなければ丸めた小数 (0、-0)。それも収まらなければ null (# で埋める)。
     */
    public static String generalFit(double v, int maxChars) {
        if (maxChars < 1 || Double.isNaN(v) || Double.isInfinite(v)) return null;
        if (v == 0) return "0";
        double a = Math.abs(v);
        String dec = null;
        boolean decOk = false;
        if (a < 1e11) {
            int intDigits = a >= 1 ? (int) Math.floor(Math.log10(a)) + 1 : 1;
            int avail = maxChars - intDigits - (v < 0 ? 1 : 0) - 1;
            int d = Math.max(0, Math.min(Math.min(avail, 30), 10 - intDigits + (a < 1 ? leadingZeros(a) : 0)));
            java.math.BigDecimal bd = new java.math.BigDecimal(v).setScale(d, java.math.RoundingMode.HALF_UP).round(new java.math.MathContext(15));
            String t = bd.signum() == 0 ? (v < 0 ? "-0" : "0") : bd.stripTrailingZeros().toPlainString();
            if (t.length() <= maxChars) {
                dec = t;
                double r = Math.abs(bd.doubleValue());
                decOk = a >= 1 || r > 0 && Math.floor(Math.log10(r) + 1e-12) == Math.floor(Math.log10(a) + 1e-12);
            }
        }
        String sci = null;
        for (int digits = Math.max(0, maxChars - 5 - (v < 0 ? 1 : 0)); digits >= 0 && sci == null; digits--) {
            String m = String.format(Locale.ROOT, "%." + digits + "E", v);
            int e = m.indexOf('E');
            String mant = m.substring(0, e);
            if (mant.contains(".")) mant = mant.replaceAll("0+$", "").replaceAll("\\.$", "");
            int exp = Integer.parseInt(m.substring(e + 1));
            String t = mant + "E" + (exp < 0 ? "-" : "+") + String.format("%02d", Math.abs(exp));
            if (t.length() <= maxChars) sci = t;
        }
        if (dec != null && decOk && (a >= 1 || sci == null || sig(dec) >= sig(sci))) return dec;
        if (sci != null) return sci;
        return dec;
    }

    /** 数字の文字列の有効数字の桁数 (指数の部分を除く)。 */
    private static int sig(String t) {
        int e = t.indexOf('E');
        String d = (e >= 0 ? t.substring(0, e) : t).replaceAll("[^0-9]", "").replaceFirst("^0+", "");
        return d.length();
    }

    /** 1 未満の数の、小数点の後に続く 0 の数 (0.001 → 2)。 */
    private static int leadingZeros(double a) {
        int n = 0;
        while (a < 0.1 && n < 20) { a *= 10; n++; }
        return n;
    }
}
