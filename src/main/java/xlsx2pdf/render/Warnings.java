package xlsx2pdf.render;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * 変換中の警告 (描けなかったセル・図形・画像など)。変換は止めずに続け、終わったら一覧を返す。
 * 変換は 1 つのスレッドで進むので、スレッドごとに集める。
 */
public final class Warnings {
    private static final ThreadLocal<Warnings> CURRENT = new ThreadLocal<>();
    /** 記録する件数の上限 (同じ内容は 1 件にまとめる)。 */
    static final int MAX = 1000;

    private final Set<String> items = new LinkedHashSet<>();
    private int dropped;
    private String sheet;

    private Warnings() {}

    /** このスレッドで集め始める。 */
    public static Warnings start() {
        Warnings w = new Warnings();
        CURRENT.set(w);
        return w;
    }

    /** このスレッドで集めるのをやめる。 */
    public static void stop() {
        CURRENT.remove();
    }

    /** これから処理するシートの名前 (警告の先頭に付ける)。 */
    public static void sheet(String name) {
        Warnings w = CURRENT.get();
        if (w != null) w.sheet = name;
    }

    /**
     * 警告を記録する。
     *
     * @param what 描けなかったもの (例: 「セル B3 の文字を描けませんでした」)
     * @param e    原因 (なければ null)
     */
    private static final boolean TRACE = System.getProperty("xlsx2pdf.trace") != null;

    public static void add(String what, Throwable e) {
        if (e != null && TRACE) e.printStackTrace();   // 開発用: -Dxlsx2pdf.trace で例外の場所を出す
        Warnings w = CURRENT.get();
        if (w == null) return;
        StringBuilder sb = new StringBuilder();
        if (w.sheet != null) sb.append("シート「").append(w.sheet).append("」: ");
        sb.append(what);
        if (e != null) {
            sb.append(" (").append(e.getClass().getSimpleName());
            if (e.getMessage() != null) sb.append(": ").append(e.getMessage().strip().replaceAll("\\s+", " "));
            sb.append(")");
        }
        String s = sb.toString();
        if (w.items.contains(s)) return;
        if (w.items.size() >= MAX) w.dropped++;
        else w.items.add(s);
    }

    /** 記録した警告の一覧。 */
    public List<String> list() {
        List<String> out = new ArrayList<>(items);
        if (dropped > 0) out.add("ほかに " + dropped + " 件の警告があります");
        return List.copyOf(out);
    }
}
