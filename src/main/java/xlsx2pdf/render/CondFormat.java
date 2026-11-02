package xlsx2pdf.render;

import java.awt.Color;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ConditionType;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.DataBarFormatting;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** セルの条件付き書式 (塗りつぶし、フォント、罫線、カラースケール、データバー) を評価する。 */
public final class CondFormat {
    /** 1 つのセルに適用する条件付き書式 (null の項目は元の書式のまま)。 */
    public static final class Result {
        /** 塗りつぶし色・文字色・データバーの色。 */
        public Color fill, fontColor, dataBar;
        /** 太字・斜体・取り消し線。 */
        public Boolean bold, italic, strike;
        /** データバーの長さ (セル幅に対する割合)。 */
        public double dataBarFraction;
        /** 値を表示しない (データバーのみ)。 */
        public boolean hideText;
        /** 罫線 (上・右・下・左)。 */
        public Edge[] borders;
        /** 表示形式 (なければ null)。 */
        public String numFmt;
    }

    /**
     * 条件付き書式の罫線 1 辺。
     *
     * @param style 線の種類
     * @param color 線の色
     */
    public record Edge(BorderStyle style, Color color) {}

    private final RenderContext ctx;
    private final XSSFSheet sheet;
    private ConditionalFormattingEvaluator eval;
    /** 数式の評価器 (FILTER の規則の数式を評価する)。 */
    private org.apache.poi.ss.formula.BaseFormulaEvaluator fe;
    /** シートに条件付き書式があるか。 */
    private final boolean any;
    /** セル (行 × 16384 + 列) ごとの評価結果。 */
    private final Map<Long, Result> cache = new HashMap<>();
    /** 規則ごとの、適用範囲の数値 (昇順)。 */
    private final Map<EvaluationConditionalFormatRule, double[]> rangeValues = new HashMap<>();

    /** シートの条件付き書式を評価する準備をする。 */
    public CondFormat(RenderContext ctx, XSSFSheet sheet) {
        this.ctx = ctx;
        this.sheet = sheet;
        this.any = sheet.getSheetConditionalFormatting().getNumConditionalFormattings() > 0;
        if (any) {
            try {
                fe = (org.apache.poi.ss.formula.BaseFormulaEvaluator) ctx.wb.getCreationHelper().createFormulaEvaluator();
                eval = new ConditionalFormattingEvaluator(ctx.wb, (org.apache.poi.ss.formula.WorkbookEvaluatorProvider) fe);
            } catch (RuntimeException e) {
                eval = null;
            }
        }
    }

    /**
     * セルに適用する条件付き書式を返す。
     *
     * @return 適用する書式。該当しなければ null
     */
    public Result eval(int r, int c) {
        if (!any || eval == null) return null;
        long key = (long) r * 16384 + c;
        if (cache.containsKey(key)) return cache.get(key);
        Result res = null;
        try {
            List<EvaluationConditionalFormatRule> rules =
                    eval.getConditionalFormattingForCell(new CellReference(sheet.getSheetName(), r, c, false, false));
            rules = withFormulaFilters(rules, r, c);
            if (!rules.isEmpty()) res = apply(rules, r, c);
        } catch (RuntimeException | StackOverflowError e) {
            res = null;
        }
        cache.put(key, res);
        return res;
    }

    /** シートの全規則 (formula を持つ「空白を含む」などの規則の判定用)。 */
    private List<EvaluationConditionalFormatRule> sheetRules;

    /**
     * 「空白を含む」「エラーを含む」「文字列を含む」などの規則 (FILTER) は、Excel は保存された数式で判定する。POI はセル自身の値で
     * 判定するため、数式がセル自身以外を指すときに食い違う。そのような規則は POI の判定を捨て、数式を評価し直す。
     */
    private List<EvaluationConditionalFormatRule> withFormulaFilters(List<EvaluationConditionalFormatRule> rules, int r, int c) {
        if (sheetRules == null) sheetRules = eval.getFormatRulesForSheet(sheet);
        List<EvaluationConditionalFormatRule> out = new ArrayList<>();
        for (EvaluationConditionalFormatRule er : rules) if (!formulaFilter(er)) out.add(er);
        boolean added = false;
        for (EvaluationConditionalFormatRule er : sheetRules) {
            if (emptyStringRule(er)) {
                // 「セルの値 = ""」(≠ "") は、空のセルも "" と等しいとみなす
                boolean inRange = false;
                for (CellRangeAddress a : er.getRegions()) if (a.isInRange(r, c)) { inRange = true; break; }
                if (!inRange || out.contains(er)) continue;
                boolean empty = isEmptyValue(r, c);
                boolean eq = er.getRule().getComparisonOperation() == org.apache.poi.ss.usermodel.ComparisonOperator.EQUAL;
                if (empty == eq) { out.add(er); added = true; }
                continue;
            }
            if (!formulaFilter(er)) continue;
            CellRangeAddress region = null;
            for (CellRangeAddress a : er.getRegions()) if (a.isInRange(r, c)) { region = a; break; }
            if (region == null) continue;
            // 相対参照の基準は、適用範囲全体を囲む範囲の左上 (Excel は複数の範囲でも 1 つの基準で数式をずらす。
            // Excel が書く「空白を含む」の数式 LEN(TRIM(K29))=0 の K29 は、適用範囲 Q29:S31 … K41:M43 を囲む範囲の左上)
            region = bounds(er.getRegions());
            try {
                // 相対参照を自前でずらしてから評価する (POI は範囲の参照 (SUM(AD12:AG12) など) をずらさない)
                String f = shiftFormula(er.getFormula1(), r - region.getFirstRow(), c - region.getFirstColumn());
                org.apache.poi.ss.formula.eval.ValueEval v = fe._getWorkbookEvaluator().evaluate(f,
                        new CellReference(sheet.getSheetName(), r, c, false, false), new CellRangeAddress(r, r, c, c));
                boolean hit = v instanceof org.apache.poi.ss.formula.eval.BoolEval b ? b.getBooleanValue()
                        : v instanceof org.apache.poi.ss.formula.eval.NumberEval n && n.getNumberValue() != 0;
                if (hit) { out.add(er); added = true; }
            } catch (RuntimeException e) {
                // 評価できない数式は該当しないものとする
            }
        }
        if (added) java.util.Collections.sort(out);
        return out;
    }

    /** 「セルの値 = ""」または「≠ ""」の規則か。 */
    private static boolean emptyStringRule(EvaluationConditionalFormatRule er) {
        if (er.getType() != ConditionType.CELL_VALUE_IS) return false;
        byte op = er.getRule().getComparisonOperation();
        return (op == org.apache.poi.ss.usermodel.ComparisonOperator.EQUAL || op == org.apache.poi.ss.usermodel.ComparisonOperator.NOT_EQUAL)
                && "\"\"".equals(er.getFormula1() == null ? null : er.getFormula1().trim());
    }

    /** セルが空 (セルなし・空白・空の文字列) か。 */
    private boolean isEmptyValue(int r, int c) {
        Row row = sheet.getRow(r);
        Cell cell = row == null ? null : row.getCell(c);
        if (cell == null) return true;
        CellType t = cell.getCellType() == CellType.FORMULA ? cell.getCachedFormulaResultType() : cell.getCellType();
        return t == CellType.BLANK || (t == CellType.STRING && cell.getStringCellValue().isEmpty());
    }

    /** 範囲をすべて囲む範囲。 */
    private static CellRangeAddress bounds(CellRangeAddress[] rs) {
        int r0 = Integer.MAX_VALUE, c0 = Integer.MAX_VALUE, r1 = -1, c1 = -1;
        for (CellRangeAddress a : rs) {
            r0 = Math.min(r0, a.getFirstRow()); c0 = Math.min(c0, a.getFirstColumn());
            r1 = Math.max(r1, a.getLastRow()); c1 = Math.max(c1, a.getLastColumn());
        }
        return new CellRangeAddress(r0, r1, c0, c1);
    }

    /** 数式の相対参照 (セル・範囲) を dr 行・dc 列ずらした数式。 */
    private String shiftFormula(String formula, int dr, int dc) {
        if (dr == 0 && dc == 0) return formula;
        var ewb = org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook.create(ctx.wb);
        org.apache.poi.ss.formula.ptg.Ptg[] ptgs = org.apache.poi.ss.formula.FormulaParser.parse(formula, ewb,
                org.apache.poi.ss.formula.FormulaType.CELL, ctx.wb.getSheetIndex(sheet));
        for (org.apache.poi.ss.formula.ptg.Ptg p : ptgs) {
            if (p instanceof org.apache.poi.ss.formula.ptg.RefPtgBase rp) {
                if (rp.isRowRelative()) rp.setRow(rp.getRow() + dr);
                if (rp.isColRelative()) rp.setColumn(rp.getColumn() + dc);
            } else if (p instanceof org.apache.poi.ss.formula.ptg.AreaPtgBase ap) {
                if (ap.isFirstRowRelative()) ap.setFirstRow(ap.getFirstRow() + dr);
                if (ap.isLastRowRelative()) ap.setLastRow(ap.getLastRow() + dr);
                if (ap.isFirstColRelative()) ap.setFirstColumn(ap.getFirstColumn() + dc);
                if (ap.isLastColRelative()) ap.setLastColumn(ap.getLastColumn() + dc);
            }
        }
        return org.apache.poi.ss.formula.FormulaRenderer.toFormulaString(ewb, ptgs);
    }

    /** 数式で判定し直す規則か (FILTER: 空白・エラー・文字列を含む等、FORMULA: 数式)。 */
    private static boolean formulaFilter(EvaluationConditionalFormatRule er) {
        if (er.getType() != ConditionType.FILTER && !(er.getType() == ConditionType.FORMULA)) return false;
        String f = er.getFormula1();
        return f != null && !f.isBlank();
    }

    /** POI の色を AWT の色にする (テーマ色を解決する)。 */
    private Color color(org.apache.poi.ss.usermodel.Color c) {
        if (c instanceof XSSFColor xc) {
            if (xc.isThemed() && ctx.wb.getStylesSource().getTheme() != null) {
                ctx.wb.getStylesSource().getTheme().inheritFromThemeAsRequired(xc);
            }
            return ctx.theme.color(xc);
        }
        return null;
    }

    /** 該当した規則を優先順に適用する (先に決まった項目は上書きしない。stopIfTrue で打ち切る)。 */
    private Result apply(List<EvaluationConditionalFormatRule> rules, int r, int c) {
        Result res = new Result();
        boolean hit = false;
        for (EvaluationConditionalFormatRule er : rules) {
            var rule = er.getRule();
            ConditionType t = er.getType();
            if (t == ConditionType.COLOR_SCALE) {
                if (res.fill == null) {
                    Color col = colorScale(er, rule.getColorScaleFormatting(), r, c);
                    if (col != null) { res.fill = col; hit = true; }
                }
                continue;
            }
            if (t == ConditionType.DATA_BAR) {
                if (res.dataBar == null) {
                    DataBarFormatting db = rule.getDataBarFormatting();
                    Double v = value(r, c);
                    double[] mm = minMax(er);
                    if (v != null && mm != null && db != null) {
                        double lo = threshold(db.getMinThreshold(), mm, er, true);
                        double hi = threshold(db.getMaxThreshold(), mm, er, false);
                        double frac = hi > lo ? (v - lo) / (hi - lo) : 1;
                        double minW = db.getWidthMin() / 100.0, maxW = db.getWidthMax() == 0 ? 0.9 : db.getWidthMax() / 100.0;
                        res.dataBarFraction = minW + Math.max(0, Math.min(1, frac)) * (maxW - minW);
                        res.dataBar = color(db.getColor());
                        if (res.dataBar == null) res.dataBar = new Color(0x638EC6);
                        res.hideText = db.isIconOnly();
                        hit = true;
                    }
                }
                continue;
            }
            if (t == ConditionType.ICON_SET) continue;
            if (res.numFmt == null && rule.getNumberFormat() != null && rule.getNumberFormat().getFormat() != null) {
                res.numFmt = rule.getNumberFormat().getFormat();
                hit = true;
            }
            PatternFormatting pf = rule.getPatternFormatting();
            if (pf != null && res.fill == null) {
                Color col = color(pf.getFillBackgroundColorColor());
                if (col == null) col = color(pf.getFillForegroundColorColor());
                if (col != null) { res.fill = col; hit = true; }
            }
            FontFormatting ff = rule.getFontFormatting();
            if (ff != null) {
                if (res.fontColor == null) {
                    Color col = color(ff.getFontColor());
                    if (col != null) { res.fontColor = col; hit = true; }
                }
                if (res.bold == null && isSet(ff, "b")) { res.bold = ff.isBold(); hit = true; }
                if (res.italic == null && isSet(ff, "i")) { res.italic = ff.isItalic(); hit = true; }
                if (res.strike == null && isSet(ff, "strike")) { res.strike = ff.isStruckout(); hit = true; }
            }
            BorderFormatting bf = rule.getBorderFormatting();
            if (bf != null) {
                if (res.borders == null) res.borders = new Edge[4];
                setB(res, 0, bf.getBorderTop(), bf.getTopBorderColorColor());
                setB(res, 1, bf.getBorderRight(), bf.getRightBorderColorColor());
                setB(res, 2, bf.getBorderBottom(), bf.getBottomBorderColorColor());
                setB(res, 3, bf.getBorderLeft(), bf.getLeftBorderColorColor());
                hit = true;
            }
            if (rule.getStopIfTrue()) break;
        }
        return hit ? res : null;
    }

    /** フォント書式で b / i / strike が指定されているか (POI が公開していないため内部の CTFont を見る)。 */
    private static boolean isSet(FontFormatting ff, String what) {
        if (ff instanceof org.apache.poi.xssf.usermodel.XSSFFontFormatting x) {
            try {
                java.lang.reflect.Field f = x.getClass().getDeclaredField("_font");
                f.setAccessible(true);
                var ct = (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont) f.get(x);
                return switch (what) {
                    case "b" -> ct.sizeOfBArray() > 0;
                    case "i" -> ct.sizeOfIArray() > 0;
                    default -> ct.sizeOfStrikeArray() > 0;
                };
            } catch (ReflectiveOperationException | RuntimeException e) {
                return false;
            }
        }
        return false;
    }

    /** 罫線の 1 辺を設定する (未設定の辺だけ)。 */
    private void setB(Result res, int side, BorderStyle bs, org.apache.poi.ss.usermodel.Color col) {
        if (res.borders[side] != null || bs == null || bs == BorderStyle.NONE) return;
        Color c = color(col);
        res.borders[side] = new Edge(bs, c == null ? Color.BLACK : c);
    }

    /** セルの数値 (数式はキャッシュされた結果)。数値でなければ null。 */
    private Double value(int r, int c) {
        Row row = sheet.getRow(r);
        if (row == null) return null;
        Cell cell = row.getCell(c);
        if (cell == null) return null;
        CellType t = cell.getCellType();
        if (t == CellType.FORMULA) t = cell.getCachedFormulaResultType();
        return t == CellType.NUMERIC ? cell.getNumericCellValue() : null;
    }

    /** 規則の適用範囲にある数値 (昇順)。 */
    private double[] values(EvaluationConditionalFormatRule er) {
        return rangeValues.computeIfAbsent(er, k -> {
            List<Double> vals = new ArrayList<>();
            for (CellRangeAddress a : er.getRegions()) {
                for (int r = a.getFirstRow(); r <= Math.min(a.getLastRow(), sheet.getLastRowNum()); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;
                    for (int c = a.getFirstColumn(); c <= Math.min(a.getLastColumn(), Math.max(0, row.getLastCellNum())); c++) {
                        Double v = value(r, c);
                        if (v != null) vals.add(v);
                    }
                }
            }
            double[] arr = vals.stream().mapToDouble(Double::doubleValue).toArray();
            Arrays.sort(arr);
            return arr;
        });
    }

    /** 規則の適用範囲の最小値と最大値 (数値がなければ null)。 */
    private double[] minMax(EvaluationConditionalFormatRule er) {
        double[] v = values(er);
        if (v.length == 0) return null;
        return new double[] {v[0], v[v.length - 1]};
    }

    /** カラースケール・データバーのしきい値を値にする (最小・最大・数値・パーセント・百分位)。 */
    private double threshold(ConditionalFormattingThreshold th, double[] mm, EvaluationConditionalFormatRule er, boolean isMin) {
        if (th == null) return isMin ? mm[0] : mm[1];
        Double val = th.getValue();
        if (val == null && th.getFormula() != null) {
            try { val = Double.parseDouble(th.getFormula()); } catch (NumberFormatException e) { val = null; }
        }
        return switch (th.getRangeType()) {
            case MIN -> mm[0];
            case MAX -> mm[1];
            case NUMBER -> val != null ? val : (isMin ? mm[0] : mm[1]);
            case PERCENT -> mm[0] + (mm[1] - mm[0]) * (val == null ? 50 : val) / 100.0;
            case PERCENTILE -> percentile(values(er), val == null ? 50 : val);
            default -> isMin ? mm[0] : mm[1];
        };
    }

    /** 百分位 (線形補間)。 */
    private static double percentile(double[] sorted, double p) {
        if (sorted.length == 0) return 0;
        double rank = p / 100.0 * (sorted.length - 1);
        int lo = (int) Math.floor(rank), hi = (int) Math.ceil(rank);
        return sorted[lo] + (sorted[hi] - sorted[lo]) * (rank - lo);
    }

    /** カラースケールの色 (しきい値の間を線形に混ぜる)。 */
    private Color colorScale(EvaluationConditionalFormatRule er, ColorScaleFormatting cs, int r, int c) {
        if (cs == null) return null;
        Double v = value(r, c);
        double[] mm = minMax(er);
        if (v == null || mm == null) return null;
        ConditionalFormattingThreshold[] th = cs.getThresholds();
        org.apache.poi.ss.usermodel.Color[] cols = cs.getColors();
        int n = Math.min(th.length, cols.length);
        if (n < 2) return null;
        double[] pos = new double[n];
        Color[] cc = new Color[n];
        for (int i = 0; i < n; i++) {
            pos[i] = threshold(th[i], mm, er, i == 0);
            cc[i] = color(cols[i]);
            if (cc[i] == null) cc[i] = Color.WHITE;
        }
        if (v <= pos[0]) return cc[0];
        for (int i = 1; i < n; i++) {
            if (v <= pos[i]) {
                double t = pos[i] > pos[i - 1] ? (v - pos[i - 1]) / (pos[i] - pos[i - 1]) : 1;
                return CellPainter.mix(cc[i], cc[i - 1], t);
            }
        }
        return cc[n - 1];
    }
}
