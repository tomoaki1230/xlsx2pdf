package xlsx2pdf.render;

import java.awt.Color;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DateUtil;

/**
 * Excel の表示形式の処理 (数値、日付、時刻、文字列セクション)。Java の書式化ではなく Excel の動作に合わせて実装する:
 * セクションの選択、条件、色、リテラル、空白 (_x)、繰り返し (*x)、千単位の縮小、パーセント、指数、分数、和暦と曜日。
 */
public final class ExcelFormat {
    /** 繰り返し文字 (*x) の位置を示す印。 */
    public static final char FILL = '\u0001';
    /** 数字のプレースホルダーの間に埋め込んだリテラルの開始・終了の印。 */
    static final char EMBED = '\u0002', EMBED_END = '\u0003';

    /** 条件 ([&lt;=100] など) の書式。 */
    private static final Pattern CONDITION = Pattern.compile("(<=|>=|<>|<|>|=)\\s*(-?[0-9.e+]+)");
    /** 秒の小数部 (ss.00 の 0 の数)。 */
    private static final Pattern FRACTION_SECONDS = Pattern.compile("s+\\.(0+)");

    /** 書式化の結果。 */
    public static final class Out {
        /** 文字列 (繰り返し位置には {@link #FILL})。 */
        public String text;
        /** セクションの [色] (なければ null)。 */
        public Color color;
        /** 繰り返し文字 (なければ null)。 */
        public Character fill;
    }

    /** 1904 年基準の日付か。 */
    private final boolean date1904;

    /** 表示形式の処理器を作る。 */
    public ExcelFormat(boolean date1904) {
        this.date1904 = date1904;
    }

    // ------------------------------------------------------------ セクション

    /** 解析した 1 つのセクション。 */
    static final class Section {
        /** 色・条件の角括弧を除いた本体。 */
        String body;
        /** [色] の指定。 */
        Color color;
        /** 条件の演算子と値。 */
        String condOp;
        double condVal;
        boolean hasCond;
        /** [$-411] などのロケール (16 進)。 */
        String locale = "";
        /** [DBNum1]〜[DBNum3] の数字の書き方 (0 はそのまま)。 */
        int dbnum;
    }

    /** 表示形式をセクションに分ける。 */
    static List<String> split(String fmt) {
        return ValueFormatter.sections(fmt);
    }

    /** セクションを解析する (角括弧の色・条件・通貨記号・ロケール・経過時間を処理する)。 */
    static Section parseSection(String s) {
        Section sec = new Section();
        StringBuilder body = new StringBuilder();
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c == '"') {
                int j = s.indexOf('"', i + 1);
                if (j < 0) j = s.length() - 1;
                body.append(s, i, j + 1);
                i = j;
            } else if (c == '\\' && i + 1 < s.length()) {
                body.append(c).append(s.charAt(++i));
            } else if (c == '[') {
                int j = s.indexOf(']', i);
                if (j < 0) { body.append(c); continue; }
                String in = s.substring(i + 1, j);
                String low = in.toLowerCase(Locale.ROOT);
                if (low.matches("h+|m+|s+")) {
                    body.append('[').append(in).append(']'); // 経過時間
                } else if (in.startsWith("$")) {
                    int dash = in.indexOf('-');
                    String sym = dash >= 0 ? in.substring(1, dash) : in.substring(1);
                    if (dash >= 0) sec.locale = in.substring(dash + 1);
                    if (!sym.isEmpty()) body.append('"').append(sym).append('"');
                } else if (low.matches("(<=|>=|<>|<|>|=)\\s*-?[0-9.e+]+")) {
                    Matcher m = CONDITION.matcher(low);
                    if (m.matches()) {
                        sec.hasCond = true;
                        sec.condOp = m.group(1);
                        sec.condVal = Double.parseDouble(m.group(2));
                    }
                } else if (low.startsWith("dbnum")) {
                    // 漢数字 (1: 一二三・十百千、2: 壱弐参・拾百阡、3: 全角数字)。数字の並びを書き換える
                    try { sec.dbnum = Integer.parseInt(low.substring(5).trim()); } catch (NumberFormatException ignored) { sec.dbnum = 0; }
                } else if (low.startsWith("natnum")) {
                    // ほかの数字の書式には対応しない: 数字のまま
                } else {
                    Color col = ValueFormatter.sectionColor("[" + in + "]");
                    if (col != null) sec.color = col;
                }
                i = j;
            } else {
                body.append(c);
            }
        }
        sec.body = body.toString();
        return sec;
    }

    /** 値がセクションの条件を満たすか。 */
    static boolean test(Section s, double v) {
        return switch (s.condOp) {
            case "<" -> v < s.condVal;
            case "<=" -> v <= s.condVal;
            case ">" -> v > s.condVal;
            case ">=" -> v >= s.condVal;
            case "=" -> v == s.condVal;
            default -> v != s.condVal;
        };
    }

    /**
     * 数値を表示形式で書式化する。
     *
     * @param fmt 表示形式 (null・空は General)
     */
    public Out format(double v, String fmt) {
        Out out = new Out();
        if (fmt == null || fmt.isEmpty()) fmt = "General";
        List<String> raw = split(fmt);
        List<Section> secs = new ArrayList<>();
        // @ を含むセクションは、位置にかかわらず文字列のセクション
        for (String r : raw) if (raw.size() == 1 || !containsUnquoted(r, '@')) secs.add(parseSection(r));
        if (secs.isEmpty()) secs.add(parseSection("General"));
        Section sec;
        boolean neg = v < 0;
        boolean useAbs = false;
        boolean anyCond = secs.stream().anyMatch(s -> s.hasCond);
        if (anyCond) {
            // 条件付き: 条件を満たす最初のセクション、なければ次の条件なしのセクション
            sec = null;
            int i = 0;
            for (; i < Math.min(2, secs.size()); i++) {
                Section s = secs.get(i);
                if (s.hasCond ? test(s, v) : (i == 0 ? v >= 0 : true)) { sec = s; break; }
            }
            if (sec == null) {
                // 条件付きの表示形式の既定のセクションは、符号付きで値を表示する
                sec = secs.size() > 2 ? secs.get(2) : secs.get(secs.size() - 1);
            }
            // Excel の動き: 条件が負の値だけを指す (< で 0 以下、= で負の値) ときだけ符号を付けない。[<5]・[<=0] は符号が残る
            useAbs = sec.hasCond && neg && (sec.condOp.equals("<") && sec.condVal <= 0 || sec.condOp.equals("=") && sec.condVal < 0);
        } else if (secs.size() == 1) {
            sec = secs.get(0);
        } else if (v > 0 || (v == 0 && secs.size() < 3)) {
            sec = secs.get(0);
        } else if (v < 0) {
            sec = secs.get(1);
            useAbs = true;
        } else {
            sec = secs.get(2);
        }
        out.color = sec.color;
        String body = sec.body;
        if (sec.hasCond && secs.size() >= 2 && neg && sec == secs.get(1)) useAbs = true;
        // 条件が最初の区画だけにある 2 区画の書式 ([>0]0;0): 2 番目の区画に来た負の値は、最初の条件の補集合が非正のときだけ符号を付けない (Excel で確認。[>5]0;0 は -5、[>0]0;0 は 5)
        if (neg && !useAbs && secs.size() == 2 && sec == secs.get(1) && !sec.hasCond && secs.get(0).hasCond) {
            Section f = secs.get(0);
            if ((f.condOp.equals(">") || f.condOp.equals(">=")) && f.condVal <= 0 || f.condOp.equals("<") && f.condVal < 0) useAbs = true;
        }
        double val = useAbs ? Math.abs(v) : v;
        // [$-F800] は OS の長い日付形式、[$-F400] は OS の時刻形式で表示し、後ろの書式 (他のアプリ用の代替) は使わない。
        // 日本語の Windows の既定は「2026年10月21日」と「13:05:09」
        if (sec.locale.equalsIgnoreCase("F800")) body = "yyyy\"年\"m\"月\"d\"日\"";
        else if (sec.locale.equalsIgnoreCase("F400")) body = "h:mm:ss";
        if (isDateFormat(body)) {
            out.text = formatDate(val, body, sec.locale);
        } else if (body.trim().isEmpty() && !body.contains("\"")) {
            out.text = body.isEmpty() ? "" : body;
        } else {
            out.text = formatNumber(val, body, true);
        }
        if (sec.dbnum >= 1 && sec.dbnum <= 3) out.text = dbnum(out.text, sec.dbnum, isDateFormat(body) ? 1 : isGeneral(body) ? 2 : 0);
        int f = out.text.indexOf(FILL);
        if (f >= 0 && f + 1 < out.text.length()) {
            out.fill = out.text.charAt(f + 1);
            out.text = out.text.substring(0, f + 1) + out.text.substring(f + 2);
        }
        return out;
    }

    private static final String KAN1 = "〇一二三四五六七八九", KAN2 = "〇壱弐参四伍六七八九";
    private static final String[] UNIT1 = {"", "十", "百", "千"}, UNIT2 = {"", "拾", "百", "阡"};
    private static final String[] BIG1 = {"", "万", "億", "兆"}, BIG2 = {"", "萬", "億", "兆"};

    /** 日付の書式の [DBNum] (dbnum(t, kind, 1))。 */
    static String dbnum(String t, int kind) {
        return dbnum(t, kind, 1);
    }

    /**
     * [DBNum] の数字の書き方に書き換える (Excel で確認): 3 は全角の数字。1・2 は、数値の書式では数字を 1 字ずつ漢数字にする (1234 → 一二三四)。
     * 標準 (mode 2) は桁数によらず位取りの漢数字 (1234 → 千二百三十四)。
     * 日付・時刻の書式 (mode 1) では、数字の並びが 1〜2 桁で 0 始まりでなければ位取りの漢数字 (12 → 十二、33 → 三十三、0 → 〇)、
     * それ以外 (4 桁の年・0 始まりの 01・小数点の後) は 1 字ずつ (二〇二〇、〇一)。
     */
    static String dbnum(String t, int kind, int mode) {
        StringBuilder o = new StringBuilder();
        int i = 0;
        while (i < t.length()) {
            char c = t.charAt(i);
            if (c < '0' || c > '9') { o.append(c); i++; continue; }
            int j = i;
            while (j < t.length() && t.charAt(j) >= '0' && t.charAt(j) <= '9') j++;
            String run = t.substring(i, j);
            boolean frac = i > 0 && t.charAt(i - 1) == '.';
            if (kind == 3) {
                for (char d : run.toCharArray()) o.append((char) ('０' + (d - '0')));
            } else if (mode == 0 || frac || run.length() > 16 || mode == 1 && (run.length() > 2 || run.length() > 1 && run.charAt(0) == '0')) {
                String k = kind == 1 ? KAN1 : KAN2;
                for (char d : run.toCharArray()) o.append(k.charAt(d - '0'));
            } else {
                o.append(kanjiInt(Long.parseLong(run), kind));
            }
            i = j;
        }
        return o.toString();
    }

    /** 整数の位取りの漢数字 (1: 千二百三十四・十 は「一」を省く、2: 壱阡弐百参拾四)。 */
    static String kanjiInt(long n, int kind) {
        String k = kind == 1 ? KAN1 : KAN2;
        if (n == 0) return String.valueOf(k.charAt(0));
        String[] unit = kind == 1 ? UNIT1 : UNIT2, big = kind == 1 ? BIG1 : BIG2;
        StringBuilder o = new StringBuilder();
        for (int g = 3; g >= 0; g--) {
            long grp = (n / (long) Math.pow(10000, g)) % 10000;
            if (grp == 0) continue;
            for (int p = 3; p >= 0; p--) {
                int d = (int) (grp / (long) Math.pow(10, p) % 10);
                if (d == 0) continue;
                if (!(p > 0 && d == 1 && kind == 1)) o.append(k.charAt(d));
                o.append(unit[p]);
            }
            o.append(big[g]);
        }
        return o.toString();
    }

    /** 表示形式の最初のセクションが「General」か (リテラル付きを含む)。 */
    public static boolean isGeneral(String fmt) {
        if (fmt == null || fmt.isEmpty()) return true;
        String b = parseSection(split(fmt).get(0)).body.trim();
        return b.equalsIgnoreCase("General");
    }

    /** 文字列のセクション (第 4 セクション、または @ を含むセクション) を適用する。なければ null。 */
    public Out formatText(String s, String fmt) {
        if (fmt == null) return null;
        List<String> raw = split(fmt);
        String secRaw = null;
        if (raw.size() >= 4) secRaw = raw.get(3);
        else if (raw.size() == 1 && containsUnquoted(raw.get(0), '@')) secRaw = raw.get(0);
        else {
            for (String r : raw) if (containsUnquoted(r, '@')) secRaw = r;
        }
        if (secRaw == null) return null;
        Section sec = parseSection(secRaw);
        Out out = new Out();
        out.color = sec.color;
        StringBuilder sb = new StringBuilder();
        String b = sec.body;
        for (int i = 0; i < b.length(); i++) {
            char c = b.charAt(i);
            if (c == '"') {
                int j = b.indexOf('"', i + 1);
                if (j < 0) j = b.length();
                sb.append(b, i + 1, j);
                i = j;
            } else if (c == '\\' && i + 1 < b.length()) sb.append(b.charAt(++i));
            else if (c == '@') sb.append(s);
            else if (c == '_' && i + 1 < b.length()) { sb.append(pad(b.charAt(++i))); }
            else if (c == '*' && i + 1 < b.length()) { sb.append(FILL); out.fill = b.charAt(++i); }
            else sb.append(c);
        }
        out.text = sb.toString();
        return out;
    }

    /** 引用符・エスケープの外に文字 ch があるか。 */
    static boolean containsUnquoted(String s, char ch) {
        boolean q = false;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c == '"') q = !q;
            else if (!q && c == '\\') i++;
            else if (!q && c == ch) return true;
        }
        return false;
    }

    /** "_x" の空白 (文字 x の幅)。括弧・マイナス・ピリオド・カンマは細い空白 (U+2009)、それ以外は半角空白。 */
    static String pad(char c) {
        // 「_x」: x の文字の幅の空白。描かない私用の文字 (U+E000 + x) で表し、幅は書体の x の幅
        if (c != ' ' && c < 0x100) return String.valueOf((char) (FontManager.PAD_BASE + c));
        return switch (c) {
            case ')', '(', '-', '.', ',' -> "\u2009";
            default -> " ";
        };
    }

    // ------------------------------------------------------------ 日付

    /** 日付・時刻の表示形式か (引用符などの外に y, m, d, h, s, g, e, aaa, 経過時間があるか)。 */
    static boolean isDateFormat(String body) {
        body = body.replaceAll("(?i)general", "");
        boolean q = false;
        for (int i = 0; i < body.length(); i++) {
            char c = body.charAt(i);
            if (c == '"') { q = !q; continue; }
            if (q) continue;
            if (c == '\\' || c == '_' || c == '*') { i++; continue; }
            if (c == '[') {
                int j = body.indexOf(']', i);
                if (j > 0 && body.substring(i + 1, j).toLowerCase(Locale.ROOT).matches("h+|m+|s+")) return true;
                if (j > 0) i = j;
                continue;
            }
            char lc = Character.toLowerCase(c);
            if (lc == 'y' || lc == 'd' || lc == 'h' || lc == 's' || lc == 'g') return true;
            if (lc == 'm') return true; // m は日付の表示形式にしか現れない
            if (lc == 'e' && !(i + 1 < body.length() && (body.charAt(i + 1) == '+' || body.charAt(i + 1) == '-'))) return true;
            if (lc == 'a' && i + 2 < body.length() && Character.toLowerCase(body.charAt(i + 1)) == 'a') return true;
        }
        return false;
    }

    private static final String[] MON = {"January", "February", "March", "April", "May", "June", "July", "August",
            "September", "October", "November", "December"};
    private static final String[] DOW = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"};
    private static final String[] DOW_JA = {"月", "火", "水", "木", "金", "土", "日"};

    /**
     * Excel の曜日 (0 = 月曜)。シリアル値から数える: 1900 年基準は 1 (1900/1/1) が日曜、0 が土曜。
     * Excel は 1900 年をうるう年として数えるので、1900/3/1 より前は実際の暦と 1 日ずれる (Excel に合わせる)。
     */
    int excelDow(double serial) {
        long sun = Math.floorMod((long) serial + (date1904 ? 5 : 6), 7);   // 0 = 日曜
        return (int) ((sun + 6) % 7);
    }

    /** 日付にできる最大のシリアル値 (9999/12/31。1904 年基準はその分小さい)。 */
    static final double MAX_DATE = 2958465, MAX_DATE_1904 = 2957003;

    /**
     * 日付・時刻を書式化する。
     *
     * @param v      シリアル値
     * @param locale セクションのロケール (411 なら日本語の月・曜日)
     */
    String formatDate(double v, String body, String locale) {
        if (!Double.isFinite(v)) return "#".repeat(8);   // 無限大・NaN は日付にも経過時間にもできない
        if (!DateUtil.isValidExcelDate(v) && !body.contains("[")) return "#".repeat(8);
        // 9999/12/31 を超える値は日付にできず、Excel は # で埋める (経過時間 [h]・[m]・[s] の書式は除く)
        boolean elapsed = body.toLowerCase(Locale.ROOT).matches("(?s).*\\[(h+|m+|s+)\\].*");
        // 経過時間でも、ほかの日付・時刻の要素を含む書式 (yyyy [h]:mm・[h]:mm・[mm]:ss・[h]:mm m など) は日付にできない値を # で埋める
        // ([h]・[m]・[s] だけの書式は上限がない)
        boolean datePart = body.replaceAll("\\[[^\\]]*\\]|\"[^\"]*\"|\\\\.", "").toLowerCase(Locale.ROOT).matches("(?s).*[ydehmsa].*");
        if ((!elapsed || datePart) && Math.floor(v) > (date1904 ? MAX_DATE_1904 : MAX_DATE)) return "#".repeat(8);
        // 秒の小数部の桁数
        int fracDigits = 0;
        Matcher fm = FRACTION_SECONDS.matcher(body.toLowerCase(Locale.ROOT));
        if (fm.find()) fracDigits = fm.group(1).length();
        double days = v;
        double rounded = Math.round(days * 86400.0 * Math.pow(10, fracDigits)) / (86400.0 * Math.pow(10, fracDigits));
        if (v < 0) return "#".repeat(8);
        LocalDateTime dt;
        try {
            dt = DateUtil.getLocalDateTime(Math.floor(rounded), date1904, false);
        } catch (RuntimeException e) {
            dt = null;
        }
        if (dt == null && !elapsed) return "#".repeat(8);   // 日付にできない値に、もっともらしい日付を出さない
        if (dt == null) dt = LocalDateTime.of(1900, 1, 1, 0, 0);
        // シリアル値 0 は Excel では 1900 年 1 月 0 日 (明治33年1月0日)
        boolean dayZero = !date1904 && Math.floor(rounded) == 0;
        if (dayZero) dt = LocalDateTime.of(1900, 1, 1, 0, 0);
        // シリアル値 60 は Excel では存在しない 1900 年 2 月 29 日 (実際の暦では 3 月 1 日になる)
        boolean fakeLeap = !date1904 && Math.floor(rounded) == 60;
        double secOfDay = (rounded - Math.floor(rounded)) * 86400.0;
        long totalSec = (long) Math.floor(secOfDay + 1e-6);
        int hh = (int) (totalSec / 3600), mi = (int) (totalSec % 3600 / 60), ss = (int) (totalSec % 60);
        double frac = secOfDay - totalSec;
        boolean ampm = body.toLowerCase(Locale.ROOT).contains("am/pm") || body.toLowerCase(Locale.ROOT).contains("a/p")
                || body.contains("午前/午後") || body.contains("上午/下午");
        String[] era = ValueFormatter.era(dt.toLocalDate());
        StringBuilder sb = new StringBuilder();
        String b = body;
        // 字句に分ける (先頭の 1 文字が種類: L リテラル、E 経過時間、P 午前/午後、T 日付の要素、F 秒の小数部、M 分)
        List<String> toks = new ArrayList<>();
        for (int i = 0; i < b.length(); ) {
            char c = b.charAt(i);
            if (c == '"') {
                int j = b.indexOf('"', i + 1);
                if (j < 0) j = b.length() - 1;
                toks.add("L" + b.substring(i + 1, j));
                i = j + 1;
            } else if (c == '\\' && i + 1 < b.length()) {
                toks.add("L" + b.charAt(i + 1));
                i += 2;
            } else if (c == '_' && i + 1 < b.length()) {
                toks.add("L" + pad(b.charAt(i + 1)));
                i += 2;
            } else if (c == '*' && i + 1 < b.length()) {
                toks.add("L" + FILL + b.charAt(i + 1));
                i += 2;
            } else if (c == '[') {
                int j = b.indexOf(']', i);
                if (j < 0) j = b.length();   // 閉じる ] がない書式は最後までを括弧の中とみなす
                toks.add("E" + b.substring(i + 1, j).toLowerCase(Locale.ROOT));
                i = j + 1;
            } else if (b.regionMatches(true, i, "AM/PM", 0, 5)) {
                toks.add("P5");
                i += 5;
            } else if (b.regionMatches(true, i, "A/P", 0, 3)) {
                toks.add("P3");
                i += 3;
            } else if (b.startsWith("午前/午後", i)) {
                toks.add("PJ");
                i += 5;
            } else if (b.startsWith("上午/下午", i)) {
                toks.add("PC");
                i += 5;
            } else {
                char lc = Character.toLowerCase(c);
                if ("ymdhsgea".indexOf(lc) >= 0) {
                    int j = i;
                    while (j < b.length() && Character.toLowerCase(b.charAt(j)) == lc) j++;
                    String t = b.substring(i, j).toLowerCase(Locale.ROOT);
                    if (lc == 'a' && t.length() < 3) {
                        toks.add("L" + b.substring(i, j));
                    } else {
                        toks.add("T" + t);
                    }
                    i = j;
                } else if (c == '.' && i + 1 < b.length() && b.charAt(i + 1) == '0') {
                    int j = i + 1;
                    while (j < b.length() && b.charAt(j) == '0') j++;
                    toks.add("F" + (j - i - 1));
                    i = j;
                } else {
                    toks.add("L" + c);
                    i++;
                }
            }
        }
        // h の後または s の前の m は分
        for (int k = 0; k < toks.size(); k++) {
            String t = toks.get(k);
            if (!t.startsWith("Tm") || t.length() > 3) continue;
            boolean minute = false;
            for (int p = k - 1; p >= 0; p--) {
                String q = toks.get(p);
                if (q.startsWith("T") || q.startsWith("E")) { minute = q.startsWith("Th") || q.equals("Eh") || q.startsWith("Eh"); break; }
            }
            for (int p = k + 1; p < toks.size() && !minute; p++) {
                String q = toks.get(p);
                if (q.startsWith("T") || q.startsWith("E")) { minute = q.startsWith("Ts") || q.startsWith("Es"); break; }
            }
            if (minute) toks.set(k, "M" + t.substring(1));
        }
        boolean ja = locale.equalsIgnoreCase("411") || locale.equalsIgnoreCase("ja-JP");
        for (String t : toks) {
            char k = t.charAt(0);
            String x = t.substring(1);
            switch (k) {
                case 'L' -> sb.append(x);
                case 'F' -> {
                    int n = Integer.parseInt(x);
                    long fr = Math.round(frac * Math.pow(10, n));
                    sb.append('.').append(String.format("%0" + n + "d", Math.min(fr, (long) Math.pow(10, n) - 1)));
                }
                case 'P' -> {
                    boolean pm = hh >= 12;
                    if (x.equals("5")) sb.append(pm ? "PM" : "AM");
                    else if (x.equals("3")) sb.append(pm ? "P" : "A");
                    else if (x.equals("C")) sb.append(pm ? "下午" : "上午");
                    else sb.append(pm ? "午後" : "午前");
                }
                case 'E' -> {
                    double totalH = rounded * 24;
                    if (x.startsWith("h")) sb.append(pad0((long) Math.floor(totalH + 1e-9), x.length()));
                    else if (x.startsWith("m")) sb.append(pad0((long) Math.floor(rounded * 1440 + 1e-9), x.length()));
                    else sb.append(pad0((long) Math.floor(rounded * 86400 + 1e-6), x.length()));
                }
                case 'M' -> sb.append(pad0(mi, x.length()));
                case 'T' -> {
                    switch (x.charAt(0)) {
                        case 'y' -> sb.append(x.length() <= 2 ? pad0(dt.getYear() % 100, 2) : String.valueOf(dt.getYear()));
                        case 'm' -> {
                            int mo = fakeLeap ? 2 : dt.getMonthValue();
                            switch (x.length()) {
                                case 1 -> sb.append(mo);
                                case 2 -> sb.append(pad0(mo, 2));
                                case 3 -> sb.append(ja ? mo + "月" : MON[mo - 1].substring(0, 3));
                                case 5 -> sb.append(MON[mo - 1].charAt(0));
                                default -> sb.append(ja ? mo + "月" : MON[mo - 1]);
                            }
                        }
                        case 'd' -> {
                            int d = dayZero ? 0 : fakeLeap ? 29 : dt.getDayOfMonth();
                            int w = excelDow(Math.floor(rounded));
                            switch (x.length()) {
                                case 1 -> sb.append(d);
                                case 2 -> sb.append(pad0(d, 2));
                                case 3 -> sb.append(ja ? DOW_JA[w] : DOW[w].substring(0, 3));
                                default -> sb.append(ja ? DOW_JA[w] + "曜日" : DOW[w]);
                            }
                        }
                        case 'a' -> {
                            int w = excelDow(Math.floor(rounded));
                            sb.append(x.length() >= 4 ? DOW_JA[w] + "曜日" : DOW_JA[w]);
                        }
                        case 'h' -> {
                            int h = ampm ? (hh % 12 == 0 ? 12 : hh % 12) : hh;
                            sb.append(x.length() >= 2 ? pad0(h, 2) : String.valueOf(h));
                        }
                        case 's' -> sb.append(x.length() >= 2 ? pad0(ss, 2) : String.valueOf(ss));
                        case 'g' -> sb.append(x.length() >= 3 ? era[0] : x.length() == 2 ? era[0].isEmpty() ? "" : era[0].substring(0, 1) : era[1]);
                        case 'e' -> {
                            String y = era[0].isEmpty() ? String.valueOf(dt.getYear()) : era[2];
                            sb.append(x.length() >= 2 && y.length() < 2 ? "0" + y : y);
                        }
                        default -> sb.append(x);
                    }
                }
                default -> { }
            }
        }
        return sb.toString();
    }

    /** 0 で n 桁に埋める。 */
    private static String pad0(long v, int n) {
        String s = String.valueOf(Math.abs(v));
        while (s.length() < n) s = "0" + s;
        return v < 0 ? "-" + s : s;
    }

    // ------------------------------------------------------------ 数値

    /**
     * 数値を書式化する。
     *
     * @param autoMinus 負の値にマイナス記号を付けるか
     */
    String formatNumber(double v, String body, boolean autoMinus) {
        // リテラルと数値のパターンに分ける
        StringBuilder pattern = new StringBuilder(); // 数字部分の文字 0#?.,%E+-/ と、その中の空白
        List<Object> parts = new ArrayList<>(); // 文字列はリテラル、Integer は数値の位置 (0 = 数値、1 = General)
        StringBuilder lit = new StringBuilder();
        boolean numberStarted = false, numberEnded = false;
        boolean percent = false;
        boolean general = false;
        int i = 0;
        String b = body;
        while (i < b.length()) {
            char c = b.charAt(i);
            if (c == '"' || (c == '\\' && i + 1 < b.length())) {
                String text;
                int next;
                if (c == '"') {
                    int j = b.indexOf('"', i + 1);
                    if (j < 0) j = b.length();
                    text = b.substring(i + 1, j);
                    next = j + 1;
                } else {
                    text = String.valueOf(b.charAt(i + 1));
                    next = i + 2;
                }
                // 数字のプレースホルダーの間のリテラルは数値の一部 ("#\ ###\ ##0"、"000-0000")
                if (numberStarted && !numberEnded && next < b.length() && "0#?".indexOf(b.charAt(next)) >= 0) {
                    pattern.append(EMBED).append(text).append(EMBED_END);
                } else {
                    lit.append(text);
                    if (numberStarted) numberEnded = true;
                }
                i = next;
                continue;
            }
            if (c == '_' && i + 1 < b.length()) { lit.append(pad(b.charAt(i + 1))); i += 2; if (numberStarted) numberEnded = true; continue; }
            if (c == '*' && i + 1 < b.length()) { lit.append(FILL).append(b.charAt(i + 1)); i += 2; if (numberStarted) numberEnded = true; continue; }
            if (b.regionMatches(true, i, "General", 0, 7)) {
                if (lit.length() > 0) { parts.add(lit.toString()); lit.setLength(0); }
                parts.add(1);
                general = true;
                i += 7;
                continue;
            }
            boolean numChar = "0#?.,".indexOf(c) >= 0 || (numberStarted && !numberEnded && (c == 'E' || c == 'e') && i + 1 < b.length()
                    && (b.charAt(i + 1) == '+' || b.charAt(i + 1) == '-')) || (numberStarted && !numberEnded && c == '/')
                    // 分母を固定する分数 (# ?/4) の分母の数字
                    || (numberStarted && !numberEnded && c >= '1' && c <= '9' && pattern.indexOf("/") >= 0);
            if (c == '%') { percent = true; lit.append('%'); i++; if (numberStarted) numberEnded = true; continue; }
            if (numChar && !numberEnded) {
                if (!numberStarted) {
                    if (lit.length() > 0) { parts.add(lit.toString()); lit.setLength(0); }
                    parts.add(0);
                    numberStarted = true;
                }
                if ((c == 'E' || c == 'e')) { pattern.append('E').append(b.charAt(i + 1)); i += 2; continue; }
                pattern.append(c);
                i++;
                continue;
            }
            if (numberStarted && !numberEnded && (c == '-' || c == ' ') && i + 1 < b.length() && "0#?".indexOf(b.charAt(i + 1)) >= 0
                    && (c == '-' || b.indexOf('/', i) < 0)) {
                pattern.append(EMBED).append(c).append(EMBED_END); // 例: 000-0000
                i++;
                continue;
            }
            if (numberStarted && !numberEnded && c == ' ' && i + 1 < b.length() && "0#?".indexOf(b.charAt(i + 1)) >= 0
                    && b.indexOf('/', i) > 0) {
                pattern.append(c); // 整数部付きの分数
                i++;
                continue;
            }
            if (numberStarted) numberEnded = true;
            lit.append(c);
            i++;
        }
        if (lit.length() > 0) parts.add(lit.toString());
        double val = percent ? v * 100 : v;
        String num;
        if (general) {
            num = ValueFormatter.general(Math.abs(val), 11);
        } else {
            String p = pattern.toString();
            // 末尾のカンマは 1000 で割る
            int scaleK = 0;
            while (p.endsWith(",")) { scaleK++; p = p.substring(0, p.length() - 1); }
            // 小数点の直前のカンマも 1000 で割る
            int dot = p.indexOf('.');
            while (dot > 0 && p.charAt(dot - 1) == ',') { scaleK++; p = p.substring(0, dot - 1) + p.substring(dot); dot--; }
            for (int k = 0; k < scaleK; k++) val /= 1000;
            num = p.isEmpty() ? "" : renderNumber(Math.abs(val), p);
        }
        boolean negative = v < 0 && autoMinus && !isZeroResult(num);
        StringBuilder sb = new StringBuilder();
        for (Object o : parts) {
            if (o instanceof Integer) {
                sb.append(num);
            } else {
                sb.append((String) o);
            }
        }
        // 符号は、前にリテラル・空白・繰り返しがあっても先頭に付く (-¥1,235、-$1,234.50。Excel で確認)
        if (negative) sb.insert(0, '-');
        return sb.toString();
    }

    /** 書式化した数値がゼロ (1〜9 の数字を含まない) か。 */
    private static boolean isZeroResult(String num) {
        for (char c : num.toCharArray()) if (c >= '1' && c <= '9') return false;
        return true;
    }

    /** 絶対値を "#,##0.00"、"0.0E+00"、"# ?/?" のような数字のパターンで文字列にする。 */
    static String renderNumber(double a, String p) {
        if (!Double.isFinite(a)) return "########";   // 無限大・NaN は数字にできない (日付側と同じく # で埋める)
        // Excel は値を有効数字 15 桁に丸めてから表示形式を当てる (1.4449999999999998 を 0.00 で 1.45)
        if (Double.isFinite(a) && a != 0) a = new BigDecimal(a).round(new java.math.MathContext(15)).doubleValue();
        int e = p.indexOf('E');
        if (e >= 0) return scientific(a, p.substring(0, e), p.substring(e + 1));
        int slash = p.indexOf('/');
        if (slash >= 0) return fraction(a, p);
        if (p.indexOf(EMBED) >= 0) return renderEmbedded(a, p);
        int dot = p.indexOf('.');
        String ip = dot >= 0 ? p.substring(0, dot) : p;
        String fp = dot >= 0 ? p.substring(dot + 1) : "";
        boolean thousands = ip.contains(",");
        ip = ip.replace(",", "");
        fp = fp.replace(",", "");
        int decimals = fp.length();
        BigDecimal bd = new BigDecimal(Double.toString(a)).setScale(decimals, RoundingMode.HALF_UP);
        String s = bd.toPlainString();
        String intPart = s.contains(".") ? s.substring(0, s.indexOf('.')) : s;
        String fracPart = s.contains(".") ? s.substring(s.indexOf('.') + 1) : "";
        // 整数部
        StringBuilder ib = new StringBuilder();
        String digits = intPart.equals("0") ? "" : intPart;
        int minInt = 0;
        for (char c : ip.toCharArray()) if (c == '0') minInt++;
        int qInt = 0;
        for (char c : ip.toCharArray()) if (c == '?') qInt++;
        while (digits.length() < minInt) digits = "0" + digits;
        if (thousands) {
            StringBuilder t = new StringBuilder();
            for (int k = 0; k < digits.length(); k++) {
                if (k > 0 && (digits.length() - k) % 3 == 0) t.append(',');
                t.append(digits.charAt(k));
            }
            digits = t.toString();
        }
        ib.append(digits);
        int pad = qInt - Math.max(0, digits.replace(",", "").length() - minInt);
        for (int k = 0; k < pad && k < qInt; k++) ib.insert(0, ' ');
        // 小数部: 0 = そのまま、# = 末尾の 0 を省く、? = 空白
        StringBuilder fb = new StringBuilder();
        if (decimals > 0) {
            char[] fc = fp.toCharArray();
            char[] fd = fracPart.toCharArray();
            int last = fd.length;
            for (int k = fd.length - 1; k >= 0; k--) {
                if (fd[k] == '0' && fc[k] != '0') last = k; else break;
            }
            for (int k = 0; k < fd.length; k++) {
                if (k < last) fb.append(fd[k]);
                else if (fc[k] == '?') fb.append(' ');
            }
        }
        String res = ib.toString();
        if (dot >= 0) res += "." + fb;
        return res;
    }

    /** プレースホルダーの間にリテラルのある整数のパターン: 数字は右から埋める。 */
    static String renderEmbedded(double a, String p) {
        int dot = p.indexOf('.');
        String ip = dot >= 0 ? p.substring(0, dot) : p;
        String rest = dot >= 0 ? p.substring(dot) : "";
        // 要素: プレースホルダー (P) またはリテラル (L)
        List<String> items = new ArrayList<>();
        boolean thousands = false;
        for (int i = 0; i < ip.length(); i++) {
            char c = ip.charAt(i);
            if (c == EMBED) {
                int j = ip.indexOf(EMBED_END, i);
                items.add("L" + ip.substring(i + 1, j));
                i = j;
            } else if (c == ',') {
                thousands = true;
            } else {
                items.add("P" + c);
            }
        }
        String plainFrac = rest.replace(String.valueOf(EMBED), "").replace(String.valueOf(EMBED_END), "");
        int decimals = 0;
        for (char c : plainFrac.toCharArray()) if (c == '0' || c == '#' || c == '?') decimals++;
        BigDecimal bd = new BigDecimal(Double.toString(a)).setScale(decimals, RoundingMode.HALF_UP);
        String sval = bd.toPlainString();
        String intDigits = sval.contains(".") ? sval.substring(0, sval.indexOf('.')) : sval;
        if (intDigits.equals("0")) intDigits = "";
        if (thousands) {
            StringBuilder t = new StringBuilder();
            for (int k = 0; k < intDigits.length(); k++) {
                if (k > 0 && (intDigits.length() - k) % 3 == 0) t.append(',');
                t.append(intDigits.charAt(k));
            }
            intDigits = t.toString();
        }
        StringBuilder out = new StringBuilder();
        int di = intDigits.length() - 1;
        int firstPlaceholder = -1;
        for (int k = 0; k < items.size(); k++) if (items.get(k).startsWith("P")) { firstPlaceholder = k; break; }
        for (int k = items.size() - 1; k >= 0; k--) {
            String it = items.get(k);
            if (it.startsWith("L")) { out.insert(0, it.substring(1)); continue; }
            char ph = it.charAt(1);
            if (k == firstPlaceholder) {
                // 最も左のプレースホルダーが残りの数字をすべて受け取る
                if (di >= 0) { out.insert(0, intDigits.substring(0, di + 1)); di = -1; }
                else if (ph == '0') out.insert(0, '0');
                else if (ph == '?') out.insert(0, ' ');
            } else if (di >= 0) {
                out.insert(0, intDigits.charAt(di--));
            } else if (ph == '0') {
                out.insert(0, '0');
            } else if (ph == '?') {
                out.insert(0, ' ');
            }
        }
        if (!rest.isEmpty()) {
            String frac = renderNumber(a - Math.floor(a), "0" + plainFrac);
            out.append(frac.substring(frac.indexOf('.')));
        }
        return out.toString();
    }

    /** 指数表示 (仮数部のパターン、E の後の符号と桁)。 */
    static String scientific(double a, String mant, String expPart) {
        char sign = expPart.charAt(0);
        String expDigits = expPart.substring(1);
        int dot = mant.indexOf('.');
        int intDigits = (dot >= 0 ? mant.substring(0, dot) : mant).replace(",", "").length();
        int decimals = dot >= 0 ? mant.substring(dot + 1).length() : 0;
        int exp = a == 0 ? 0 : (int) Math.floor(Math.log10(a));
        if (intDigits > 1) exp = (int) Math.floor((double) exp / intDigits) * intDigits;
        double m = a == 0 ? 0 : a / Math.pow(10, exp);
        BigDecimal bd = new BigDecimal(m).setScale(decimals, RoundingMode.HALF_UP);
        if (bd.compareTo(BigDecimal.TEN.pow(Math.max(1, intDigits))) >= 0) {
            exp += Math.max(1, intDigits);
            bd = new BigDecimal(a / Math.pow(10, exp)).setScale(decimals, RoundingMode.HALF_UP);
        }
        String ms = renderNumber(bd.doubleValue(), mant);
        String es = String.valueOf(Math.abs(exp));
        while (es.length() < expDigits.length()) es = "0" + es;
        return ms + "E" + (exp < 0 ? "-" : sign == '+' ? "+" : "") + es;
    }

    /** 分数の ? の空白: 数字 1 文字の幅 (U+2007)。 */
    private static final String FIG = "\u2007";

    /** 文字列を幅 w まで空白で埋める (left なら左に埋めて右詰め)。 */
    private static String pad(String v, int w, boolean left) {
        if (v.length() >= w) return v;
        String sp = FIG.repeat(w - v.length());
        return left ? sp + v : v + sp;
    }

    /** 分数 (分母の固定 "/4"、または桁数 "/??" に対する最良近似)。 */
    static String fraction(double a, String p) {
        String[] parts = p.split("/");
        String left = parts[0].trim();
        String den = parts.length > 1 ? parts[1].trim() : "?";   // "# ?/" のように分母のない書式は 1 桁の分母
        boolean hasInt = left.contains(" ");
        long whole = hasInt ? (long) Math.floor(a) : 0;
        double f = a - whole;
        long bestN = 0, bestD = 1;
        if (den.matches("\\d+")) {
            bestD = Long.parseLong(den);
            bestN = Math.round(f * bestD);
        } else {
            // Excel は連分数の近似分数 (convergent) のうち、分母が桁数に収まる最後のものを使う
            // (12.1 を # ?/? で 12、# ??/?? で 12 1/10。最も近い 1/9 は選ばない)
            long maxD = (long) Math.pow(10, den.length()) - 1;
            long h0 = 0, h1 = 1, k0 = 1, k1 = 0;
            double x = f;
            bestN = 0; bestD = 1;
            for (int it = 0; it < 64; it++) {
                long q = (long) Math.floor(x);
                long h2 = q * h1 + h0, k2 = q * k1 + k0;
                if (k2 > maxD) break;
                bestN = h2; bestD = k2;
                h0 = h1; h1 = h2; k0 = k1; k1 = k2;
                double r = x - q;
                if (r < 1e-12 || Math.abs(f - (double) h2 / k2) < 1e-12) break;
                x = 1 / r;
            }
        }
        if (bestN == bestD && hasInt) { whole++; bestN = 0; }
        // 分子・分母の桁の指定 (? は桁がなければ空白)。分子は右に、分母は左に詰めて、指定の幅まで空白で埋める
        String numPat = hasInt ? left.substring(left.lastIndexOf(' ') + 1) : left;
        int numW = numPat.replaceAll("[^#0?]", "").length();
        int denW = den.matches("\\d+") ? 0 : den.replaceAll("[^#0?]", "").length();
        StringBuilder sb = new StringBuilder();
        if (hasInt && (whole > 0 || left.substring(0, left.lastIndexOf(' ')).contains("0"))) sb.append(whole).append(' ');
        if (bestN == 0 && hasInt) {
            if (sb.length() == 0) sb.append('0');
            // 分数の部分がなくなっても、その幅の空白を残す (Excel は桁を揃える)
            String t = sb.toString().trim();
            // 分母を決めた形 (# ?/4) も同じ (3 は「3」の後に「 ?/4」の幅の空白)
            boolean fixedDen = den.matches("\\d+");
            return numW > 0 && (denW > 0 || fixedDen) ? t + " " + FIG.repeat(numW + 1 + (fixedDen ? den.length() : Math.max(denW, den.length()))) : t;
        }
        sb.append(pad(String.valueOf(bestN), numW, true)).append('/').append(pad(String.valueOf(bestD), denW, false));
        return sb.toString();
    }
}
