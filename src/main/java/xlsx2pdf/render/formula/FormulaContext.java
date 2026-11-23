package xlsx2pdf.render.formula;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * --recalc の関数の入口と、計算中の状態 (計算しているブック・シート・ファイルのパス)。
 * 関数の実装はブックを知らずに動くので、ブックの情報が要る関数 (CELL・ISFORMULA・AGGREGATE など) はここから読む。
 * 状態はスレッドごとに持つ (同時に別のブックを計算しても混ざらない)。計算の後は {@link #clear} で外す。
 */
public final class FormulaContext {
    private FormulaContext() {
    }

    private static final class State {
        XSSFWorkbook book;
        int sheet = -1;
        String filePath;
    }

    private static final ThreadLocal<State> STATE = ThreadLocal.withInitial(State::new);

    /** 再計算の準備: 関数を登録し、このスレッドで計算するブックを覚える。 */
    public static void register(XSSFWorkbook wb) {
        FunctionRegistry.register(wb);
        State s = new State();
        s.book = wb;
        STATE.set(s);
    }

    /** このスレッドの計算の状態を外す (ブックを持ち続けない)。 */
    public static void clear() {
        STATE.remove();
    }

    /** ブックのファイルのパスを設定する (CELL("filename") 用。分からなければ null)。 */
    public static void setFilePath(String path) {
        STATE.get().filePath = path;
    }

    /** 計算中のセルのあるシートの番号を設定する (CELL の address が他のシートの参照かを調べる)。 */
    public static void setCurrentSheet(int index) {
        STATE.get().sheet = index;
    }

    /** 再計算中のブック (なければ null)。 */
    static XSSFWorkbook workbook() {
        return STATE.get().book;
    }

    static int currentSheet() {
        return STATE.get().sheet;
    }

    /** ブックのファイルの絶対パス (分からなければ null)。 */
    static String filePath() {
        return STATE.get().filePath;
    }

    /** ブックが 1904 年の日付システムか。 */
    static boolean is1904() {
        XSSFWorkbook w = STATE.get().book;
        return w != null && w.isDate1904();
    }
}
