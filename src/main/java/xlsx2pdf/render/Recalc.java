package xlsx2pdf.render;

import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import xlsx2pdf.render.formula.FormulaContext;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;

/**
 * ブックの数式を再計算し、セルに保存された計算結果 (キャッシュ) を置き換える ({@code --recalc})。
 *
 * <p>計算は Apache POI の計算エンジンで行う。計算できない数式 (POI にない関数、データテーブル、
 * 外部参照のうち値の残っていないものなど) は、保存された結果をそのまま使う。POI が「#NAME?」を返したもの
 * (POI の知らない関数) も、保存された結果がエラーでなければ保存された結果を使う。
 */
public final class Recalc {
    private Recalc() {
    }

    /** 再計算の結果の件数。 */
    public record Result(int formulas, int updated, int kept) {
    }

    /**
     * ブックのすべてのシートの数式を再計算する。
     *
     * @param wb       ブック
     * @param deadline 時間切れの確認 (なければ null)
     * @return 件数 (数式のセル、計算した結果で置き換えたセル、保存された結果のままにしたセル)
     */
    public static Result run(XSSFWorkbook wb, Runnable deadline) {
        return run(wb, deadline, null);
    }

    /**
     * @param filePath ブックのファイルのパス (CELL("filename") 用。分からなければ null)
     */
    public static Result run(XSSFWorkbook wb, Runnable deadline, String filePath) {
        int formulas = 0, updated = 0, kept = 0;
        java.util.Map<String, java.util.List<String>> keptAt = new java.util.LinkedHashMap<>(), errorAt = new java.util.LinkedHashMap<>();
        FormulaContext.register(wb);
        try {
            FormulaContext.setFilePath(filePath);
            XSSFFormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            // 外部のブックへの参照は、ブックに残っている値 (externalLink のキャッシュ) で計算する
            ev.setIgnoreMissingWorkbooks(true);
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                XSSFSheet sheet = wb.getSheetAt(i);
                FormulaContext.setCurrentSheet(i);
                for (Row row : sheet) {
                    if (deadline != null) deadline.run();
                    for (Cell c : row) {
                        if (c.getCellType() != CellType.FORMULA) continue;
                        formulas++;
                        XSSFCell cell = (XSSFCell) c;
                        CTCell saved = (CTCell) cell.getCTCell().copy();
                        boolean savedError = saved.isSetT() && saved.getT() == org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType.E;
                        boolean ok;
                        try {
                            ev.evaluateFormulaCell(cell);
                            ok = savedError || !unknownName(cell);
                        } catch (RuntimeException | StackOverflowError e) {
                            // 計算できない数式 (循環に近い深い参照・回数の多すぎる計算を含む) は保存された結果のまま
                            ok = false;
                        }
                        if (ok) {
                            updated++;
                            // 保存された結果はエラーでないのに、計算の結果がエラー (#REF!・#VALUE! など) になったセル
                            if (!savedError && cell.getCachedFormulaResultType() == CellType.ERROR)
                                errorAt.computeIfAbsent(sheet.getSheetName(), k -> new java.util.ArrayList<>())
                                        .add(cell.getAddress().formatAsString() + " (" + errName(cell.getErrorCellValue()) + ")");
                        } else {
                            cell.getCTCell().set(saved);
                            kept++;
                            keptAt.computeIfAbsent(sheet.getSheetName(), k -> new java.util.ArrayList<>()).add(cell.getAddress().formatAsString());
                        }
                    }
                }
            }
        } finally {
            FormulaContext.clear();
        }
        Warnings.sheet(null);
        if (kept > 0)
            Warnings.add("--recalc: 計算できない数式 " + kept + " 件は、ファイルに保存された結果を使いました: " + cells(keptAt), null);
        int errors = errorAt.values().stream().mapToInt(java.util.List::size).sum();
        if (errors > 0)
            Warnings.add("--recalc: 数式 " + errors + " 件の計算の結果がエラーになりました (保存された結果はエラーではありません): " + cells(errorAt), null);
        return new Result(formulas, updated, kept);
    }

    /** 警告に出すセルの番地の上限。 */
    static final int MAX_CELLS = 20;

    /** シートごとのセルの番地を「シート「名前」 A1, B2」の形に並べる (多ければ先頭だけ)。 */
    static String cells(java.util.Map<String, java.util.List<String>> at) {
        StringBuilder sb = new StringBuilder();
        int shown = 0, total = 0;
        for (var e : at.entrySet()) {
            total += e.getValue().size();
            if (shown >= MAX_CELLS) continue;
            if (sb.length() > 0) sb.append("、");
            sb.append("シート「").append(e.getKey()).append("」 ");
            int k = 0;
            for (String a : e.getValue()) {
                if (shown >= MAX_CELLS) break;
                if (k++ > 0) sb.append(", ");
                sb.append(a);
                shown++;
            }
        }
        if (total > shown) sb.append(" ほか ").append(total - shown).append(" 件");
        return sb.toString();
    }

    /** エラーの値の表記 (#REF! など)。 */
    private static String errName(byte code) {
        try {
            return FormulaError.forInt(code).getString();
        } catch (IllegalArgumentException e) {
            return "#VALUE!";
        }
    }

    /** 計算の結果が #NAME? (POI の知らない関数・名前) か。 */
    private static boolean unknownName(XSSFCell cell) {
        if (cell.getCachedFormulaResultType() != CellType.ERROR) return false;
        return cell.getErrorCellValue() == FormulaError.NAME.getCode()
                || cell.getErrorCellValue() == (byte) ErrorEval.NAME_INVALID.getErrorCode();
    }
}
