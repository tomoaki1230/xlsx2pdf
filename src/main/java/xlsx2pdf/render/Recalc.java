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
        XSSFFormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
        // 外部のブックへの参照は、ブックに残っている値 (externalLink のキャッシュ) で計算する
        ev.setIgnoreMissingWorkbooks(true);
        int formulas = 0, updated = 0, kept = 0;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            for (Row row : sheet) {
                if (deadline != null) deadline.run();
                for (Cell c : row) {
                    if (c.getCellType() != CellType.FORMULA) continue;
                    formulas++;
                    XSSFCell cell = (XSSFCell) c;
                    CTCell saved = (CTCell) cell.getCTCell().copy();
                    boolean savedError = saved.isSetT() && saved.getT() == org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType.E;
                    try {
                        ev.evaluateFormulaCell(cell);
                        if (!savedError && unknownName(cell)) {
                            cell.getCTCell().set(saved);
                            kept++;
                        } else {
                            updated++;
                        }
                    } catch (RuntimeException e) {
                        // 計算できない数式は保存された結果のまま
                        cell.getCTCell().set(saved);
                        kept++;
                    } catch (StackOverflowError e) {
                        // 循環に近い深い参照: 保存された結果のまま
                        cell.getCTCell().set(saved);
                        kept++;
                    }
                }
            }
        }
        if (kept > 0) {
            Warnings.sheet(null);
            Warnings.add("--recalc: 計算できない数式 " + kept + " 件は、ファイルに保存された結果を使いました", null);
        }
        return new Result(formulas, updated, kept);
    }

    /** 計算の結果が #NAME? (POI の知らない関数・名前) か。 */
    private static boolean unknownName(XSSFCell cell) {
        if (cell.getCachedFormulaResultType() != CellType.ERROR) return false;
        return cell.getErrorCellValue() == FormulaError.NAME.getCode()
                || cell.getErrorCellValue() == (byte) ErrorEval.NAME_INVALID.getErrorCode();
    }
}
