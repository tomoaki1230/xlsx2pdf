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
        /** 塗りつぶし色・文字色。 */
        public Color fill, fontColor;
        /** 網掛けの塗り (patBits が 0 でなければ有効。模様は CellPainter.PATTERNS、前景色・背景色)。 */
        public long patBits;
        public Color patFg, patBg;
        /** 太字・斜体・取り消し線。 */
        public Boolean bold, italic, strike;
        /** 値を表示しない (データバーのみ)。 */
        public boolean hideText;
        /** 罫線 (上・右・下・左)。 */
        public Edge[] borders;
        /** 表示形式 (なければ null)。 */
        public String numFmt;
        /** アイコンセットのアイコン (なければ null)。 */
        IconSets.Icon icon;
        /** データバー (なければ null)。 */
        DataBars.Bar bar;
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
    /** アイコンセット・データバー (拡張の規則を含む。なければ null)。 */
    private IconSets icons;
    private DataBars bars;
    /** 拡張 (x14) の書式の規則 (数式・セルの値・文字列を含むなど。優先順位の順。なければ空)。 */
    private final List<CfExt.X14Rule> x14fmt = new ArrayList<>();
    /** セル (行 × 16384 + 列) ごとの評価結果。 */
    private final Map<Long, Result> cache = new HashMap<>();
    /** 規則ごとの、適用範囲の数値 (昇順)。 */
    private final Map<EvaluationConditionalFormatRule, double[]> rangeValues = new HashMap<>();

    /** シートの条件付き書式を評価する準備をする。 */
    public CondFormat(RenderContext ctx, XSSFSheet sheet) {
        this.ctx = ctx;
        this.sheet = sheet;
        this.any = sheet.getSheetConditionalFormatting().getNumConditionalFormattings() > 0;
        try {
            List<CfExt.X14Rule> x14 = CfExt.x14Rules(sheet);
            icons = IconSets.parse(sheet, x14);
            bars = DataBars.parse(sheet, x14, this::xmlColor);
            for (CfExt.X14Rule xr : x14) {
                String ty = xr.rule().getAttribute("type");
                if (!ty.equals("dataBar") && !ty.equals("iconSet") && !ty.equals("colorScale") && Xml.child(xr.rule(), "dxf") != null) x14fmt.add(xr);
            }
            x14fmt.sort(java.util.Comparator.comparingInt(xr -> Xml.intAttr(xr.rule(), "priority", Integer.MAX_VALUE)));
        } catch (RuntimeException e) {
            Warnings.add("アイコンセット・データバーを読めませんでした", e);
        }
        if (any) {
            try {
                fe = (org.apache.poi.ss.formula.BaseFormulaEvaluator) ctx.wb.getCreationHelper().createFormulaEvaluator();
                eval = new ConditionalFormattingEvaluator(ctx.wb, (org.apache.poi.ss.formula.WorkbookEvaluatorProvider) fe);
            } catch (RuntimeException e) {
                Warnings.add("条件付き書式を評価できません", e);
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
        if (icons == null && bars == null && x14fmt.isEmpty() && (!any || eval == null)) return null;
        long key = (long) r * 16384 + c;
        if (cache.containsKey(key)) return cache.get(key);
        Result res = null;
        if (any && eval != null) {
            try {
                List<EvaluationConditionalFormatRule> rules =
                        eval.getConditionalFormattingForCell(new CellReference(sheet.getSheetName(), r, c, false, false));
                rules = withFormulaFilters(rules, r, c);
                if (!rules.isEmpty()) res = apply(rules, r, c);
            } catch (RuntimeException | StackOverflowError e) {
                Warnings.add("セル " + new CellReference(r, c).formatAsString() + " の条件付き書式を評価できませんでした", e);
                res = null;
            }
        }
        if (!x14fmt.isEmpty()) {
            try {
                res = applyX14(res, r, c);
            } catch (RuntimeException | StackOverflowError e) {
                Warnings.add("セル " + new CellReference(r, c).formatAsString() + " の条件付き書式 (拡張) を評価できませんでした", e);
            }
        }
        if (bars != null) {
            try {
                DataBars.Bar b = bars.at(r, c, this::value, f -> formulaNumber(f, r, c));
                if (b != null) {
                    if (res == null) res = new Result();
                    res.bar = b;
                    res.hideText = b.hideText();
                }
            } catch (RuntimeException e) {
                Warnings.add("セル " + new CellReference(r, c).formatAsString() + " のデータバーを評価できませんでした", e);
            }
        }
        if (icons != null) {
            try {
                IconSets.Icon ic = icons.at(r, c, this::value, f -> formulaNumber(f, r, c));
                if (ic != null && IconPainter.known(ic.set())) {
                    if (res == null) res = new Result();
                    res.icon = ic;
                }
            } catch (RuntimeException e) {
                Warnings.add("セル " + new CellReference(r, c).formatAsString() + " のアイコンセットを評価できませんでした", e);
            }
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

    /**
     * 拡張 (x14:cfRule) の書式の規則をセルに適用する。ほかのシートを参照する数式の規則などは、Excel が拡張の側に保存する
     * (表紙!$E$9<>"✓" で全面を網掛けにする)。数式 (xm:f) は適用範囲を囲む範囲の左上を基準にずらして評価する。
     * 種類が cellIs のものはセルの値と比べ、それ以外 (expression・containsText など) は 1 つ目の数式の真偽で判定する。
     * 通常の規則で決まった項目は上書きしない。
     */
    private Result applyX14(Result res, int r, int c) {
        for (CfExt.X14Rule xr : x14fmt) {
            if (!CfExt.in(xr.ranges(), r, c)) continue;
            org.w3c.dom.Element rule = xr.rule();
            List<org.w3c.dom.Element> fs = Xml.children(rule, "f");
            if (fs.isEmpty()) continue;
            CellRangeAddress base = bounds(xr.ranges().toArray(new CellRangeAddress[0]));
            int dr = r - base.getFirstRow(), dc = c - base.getFirstColumn();
            boolean hit;
            if (rule.getAttribute("type").equals("cellIs")) {
                Double v = value(r, c), a = formulaNumber(shiftFormula(Xml.text(fs.get(0)), dr, dc), r, c);
                Double b = fs.size() > 1 ? formulaNumber(shiftFormula(Xml.text(fs.get(1)), dr, dc), r, c) : null;
                if (v == null || a == null) continue;
                hit = switch (rule.getAttribute("operator")) {
                    case "equal" -> v.doubleValue() == a;
                    case "notEqual" -> v.doubleValue() != a;
                    case "greaterThan" -> v > a;
                    case "greaterThanOrEqual" -> v >= a;
                    case "lessThan" -> v < a;
                    case "lessThanOrEqual" -> v <= a;
                    case "between" -> b != null && v >= Math.min(a, b) && v <= Math.max(a, b);
                    case "notBetween" -> b != null && (v < Math.min(a, b) || v > Math.max(a, b));
                    default -> false;
                };
            } else {
                hit = formulaTrue(shiftFormula(Xml.text(fs.get(0)), dr, dc), r, c);
            }
            if (!hit) continue;
            if (res == null) res = new Result();
            applyDxf(res, Xml.child(rule, "dxf"));
            if (Xml.boolAttr(rule, "stopIfTrue", false)) break;
        }
        return res;
    }

    /** 数式の真偽 (評価できなければ偽)。 */
    private boolean formulaTrue(String f, int r, int c) {
        try {
            if (fe == null) fe = (org.apache.poi.ss.formula.BaseFormulaEvaluator) ctx.wb.getCreationHelper().createFormulaEvaluator();
            org.apache.poi.ss.formula.eval.ValueEval v = fe._getWorkbookEvaluator().evaluate(f,
                    new CellReference(sheet.getSheetName(), r, c, false, false), new CellRangeAddress(r, r, c, c));
            v = org.apache.poi.ss.formula.eval.OperandResolver.getSingleValue(v, r, c);
            return v instanceof org.apache.poi.ss.formula.eval.BoolEval b ? b.getBooleanValue()
                    : v instanceof org.apache.poi.ss.formula.eval.NumberEval n && n.getNumberValue() != 0;
        } catch (RuntimeException | org.apache.poi.ss.formula.eval.EvaluationException e) {
            return false;
        }
    }

    /** 拡張の規則の差分書式 (x14:dxf) の塗り・フォント・罫線・表示形式を、まだ決まっていない項目に入れる。 */
    private void applyDxf(Result res, org.w3c.dom.Element dxf) {
        org.w3c.dom.Element pf = Xml.path(dxf, "fill", "patternFill");
        if (pf != null && res.fill == null && res.patBits == 0) {
            String pt = Xml.attr(pf, "patternType", "solid");
            org.w3c.dom.Element fgE = Xml.child(pf, "fgColor"), bgE = Xml.child(pf, "bgColor");
            Color fg = fgE == null ? null : Xml.boolAttr(fgE, "auto", false) ? Color.BLACK : xmlColor(fgE);
            Color bg = bgE == null ? null : Xml.boolAttr(bgE, "auto", false) ? Color.WHITE : xmlColor(bgE);
            Long bits = CellPainter.PATTERNS.get(pt);
            if (pt.equals("solid")) res.fill = bg != null ? bg : fg;
            else if (bits != null) {
                res.patBits = bits;
                res.patFg = fg != null ? fg : Color.BLACK;
                res.patBg = bg != null ? bg : Color.WHITE;
            }
        }
        org.w3c.dom.Element font = Xml.child(dxf, "font");
        if (font != null) {
            org.w3c.dom.Element col = Xml.child(font, "color");
            if (res.fontColor == null && col != null) res.fontColor = Xml.boolAttr(col, "auto", false) ? Color.BLACK : xmlColor(col);
            if (res.bold == null && Xml.child(font, "b") != null) res.bold = Xml.boolAttr(Xml.child(font, "b"), "val", true);
            if (res.italic == null && Xml.child(font, "i") != null) res.italic = Xml.boolAttr(Xml.child(font, "i"), "val", true);
            if (res.strike == null && Xml.child(font, "strike") != null) res.strike = Xml.boolAttr(Xml.child(font, "strike"), "val", true);
        }
        org.w3c.dom.Element nf = Xml.child(dxf, "numFmt");
        if (res.numFmt == null && nf != null && nf.hasAttribute("formatCode")) res.numFmt = nf.getAttribute("formatCode");
        org.w3c.dom.Element bd = Xml.child(dxf, "border");
        if (bd != null) {
            String[] sides = {"top", "right", "bottom", "left"};
            for (int i = 0; i < 4; i++) {
                org.w3c.dom.Element e = Xml.child(bd, sides[i]);
                if (e == null || !e.hasAttribute("style")) continue;
                BorderStyle bs;
                try {
                    bs = BorderStyle.valueOf(e.getAttribute("style").replaceAll("([A-Z])", "_$1").toUpperCase(java.util.Locale.ROOT));
                } catch (IllegalArgumentException ex) {
                    continue;
                }
                if (res.borders == null) res.borders = new Edge[4];
                if (res.borders[i] == null) {
                    org.w3c.dom.Element col = Xml.child(e, "color");
                    Color cc = col == null || Xml.boolAttr(col, "auto", false) ? Color.BLACK : xmlColor(col);
                    res.borders[i] = new Edge(bs, cc != null ? cc : Color.BLACK);
                }
            }
        }
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
            // データバーとアイコンセットは DataBars・IconSets で扱う
            if (t == ConditionType.DATA_BAR) continue;
            if (t == ConditionType.ICON_SET) continue;
            if (res.numFmt == null && rule.getNumberFormat() != null && rule.getNumberFormat().getFormat() != null) {
                res.numFmt = rule.getNumberFormat().getFormat();
                hit = true;
            }
            PatternFormatting pf = rule.getPatternFormatting();
            if (pf != null && res.fill == null && res.patBits == 0 && pf.getFillPattern() > 1 && pf.getFillPattern() < CellPainter.PATTERN_NAMES.length
                    && CellPainter.PATTERNS.get(CellPainter.PATTERN_NAMES[pf.getFillPattern()]) != null) {
                // 網掛けの塗り (dxf の patternType が solid 以外): セルの網掛けと同じ模様で描く
                res.patBits = CellPainter.PATTERNS.get(CellPainter.PATTERN_NAMES[pf.getFillPattern()]);
                Color fg = color(pf.getFillForegroundColorColor()), bg = color(pf.getFillBackgroundColorColor());
                res.patFg = fg != null ? fg : Color.BLACK;
                res.patBg = bg != null ? bg : Color.WHITE;
                hit = true;
            } else if (pf != null && res.fill == null && res.patBits == 0) {
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
                if (res.bold == null && isSet(er, "b")) { res.bold = ff.isBold(); hit = true; }
                if (res.italic == null && isSet(er, "i")) { res.italic = ff.isItalic(); hit = true; }
                if (res.strike == null && isSet(er, "strike")) { res.strike = ff.isStruckout(); hit = true; }
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

    /**
     * フォント書式で b / i / strike が指定されているか。POI の FontFormatting は未指定と false を区別しないので、
     * 規則の dxf (シートの cfRule の dxfId が指すスタイルの差分書式) のフォントを見る。
     */
    private boolean isSet(EvaluationConditionalFormatRule er, String what) {
        var ws = sheet.getCTWorksheet();
        int f = er.getFormattingIndex(), r = er.getRuleIndex();
        if (f < 0 || f >= ws.sizeOfConditionalFormattingArray()) return false;
        var cf = ws.getConditionalFormattingArray(f);
        if (r < 0 || r >= cf.sizeOfCfRuleArray()) return false;
        var rule = cf.getCfRuleArray(r);
        if (!rule.isSetDxfId()) return false;
        var dxf = ctx.wb.getStylesSource().getDxfAt((int) rule.getDxfId());
        if (dxf == null || !dxf.isSetFont()) return false;
        var font = dxf.getFont();
        return switch (what) {
            case "b" -> font.sizeOfBArray() > 0;
            case "i" -> font.sizeOfIArray() > 0;
            default -> font.sizeOfStrikeArray() > 0;
        };
    }

    /** 罫線の 1 辺を設定する (未設定の辺だけ)。 */
    private void setB(Result res, int side, BorderStyle bs, org.apache.poi.ss.usermodel.Color col) {
        if (res.borders[side] != null || bs == null || bs == BorderStyle.NONE) return;
        Color c = color(col);
        res.borders[side] = new Edge(bs, c == null ? Color.BLACK : c);
    }

    /** セルの数値 (数式はキャッシュされた結果)。数値でなければ null。 */
    /**
     * 色の指定 (通常の規則の CTColor、拡張の x14:fillColor などの要素) を AWT の色にする。
     *
     * @return 色。指定がなければ null
     */
    private Color xmlColor(Object o) {
        try {
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor ct;
            if (o instanceof org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor c) ct = c;
            else if (o instanceof org.w3c.dom.Element e) {
                ct = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor.Factory.newInstance();
                if (e.hasAttribute("rgb")) ct.setRgb(java.util.HexFormat.of().parseHex(e.getAttribute("rgb")));
                if (e.hasAttribute("theme")) ct.setTheme(Long.parseLong(e.getAttribute("theme")));
                if (e.hasAttribute("indexed")) ct.setIndexed(Long.parseLong(e.getAttribute("indexed")));
                if (e.hasAttribute("tint")) ct.setTint(Double.parseDouble(e.getAttribute("tint")));
                if (e.hasAttribute("auto")) ct.setAuto(Xml.boolAttr(e, "auto", false));
            } else return null;
            return color(XSSFColor.from(ct, ctx.wb.getStylesSource().getIndexedColors()));
        } catch (RuntimeException e) {
            Warnings.add("条件付き書式の色を読めませんでした", e);
            return null;
        }
    }

    /** しきい値の数式を評価した数値 (評価できなければ null)。 */
    private Double formulaNumber(String f, int r, int c) {
        if (fe == null) {
            try {
                fe = (org.apache.poi.ss.formula.BaseFormulaEvaluator) ctx.wb.getCreationHelper().createFormulaEvaluator();
            } catch (RuntimeException e) {
                return null;
            }
        }
        try {
            org.apache.poi.ss.formula.eval.ValueEval v = fe._getWorkbookEvaluator().evaluate(f,
                    new CellReference(sheet.getSheetName(), r, c, false, false), new CellRangeAddress(r, r, c, c));
            // セルの参照 ($G$1 など) は値を取り出す
            v = org.apache.poi.ss.formula.eval.OperandResolver.getSingleValue(v, r, c);
            return v instanceof org.apache.poi.ss.formula.eval.NumberEval n ? n.getNumberValue() : null;
        } catch (RuntimeException | org.apache.poi.ss.formula.eval.EvaluationException e) {
            return null;
        }
    }

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
    private double threshold(ConditionalFormattingThreshold th, double[] mm, EvaluationConditionalFormatRule er, boolean isMin, int r, int c) {
        if (th == null) return isMin ? mm[0] : mm[1];
        Double val = th.getValue();
        if (val == null && th.getFormula() != null) {
            String f = th.getFormula().trim();
            if (f.startsWith("=")) f = f.substring(1);
            try { val = Double.parseDouble(f); } catch (NumberFormatException e) { val = formulaNumber(f, r, c); } // 数式は評価する
        }
        if (th.getRangeType() == ConditionalFormattingThreshold.RangeType.FORMULA) return val != null ? val : (isMin ? mm[0] : mm[1]);
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
            pos[i] = threshold(th[i], mm, er, i == 0, r, c);
            cc[i] = color(cols[i]);
            if (cc[i] == null) cc[i] = Color.WHITE;
        }
        // 最大以上は最大の色 (全部が同じ値・1 セルだけの範囲も最大の色)
        if (v >= pos[n - 1]) return cc[n - 1];
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
