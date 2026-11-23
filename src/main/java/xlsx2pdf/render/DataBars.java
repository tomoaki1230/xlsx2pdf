package xlsx2pdf.render;

import java.awt.Color;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;
import java.util.function.Function;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfvo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataBar;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.w3c.dom.Element;

/**
 * 条件付き書式のデータバー。通常の規則 (Excel 2007 の形) と、それを置き換える Excel 2010 以降の拡張 (x14:dataBar。
 * 長さの範囲、単色・グラデーション、枠線、負の値の色、軸の位置、方向) を読む。
 */
final class DataBars {
    /**
     * セルに描く棒。位置は棒を描ける範囲 (セルの内側) に対する割合。
     *
     * @param start    棒の始まり (左端、0〜1)
     * @param end      棒の終わり (右端、0〜1)
     * @param axis     軸の位置 (0〜1。軸がなければ NaN)
     * @param fill     塗りの色
     * @param border   枠線の色 (なければ null)
     * @param gradient グラデーション (濃い側は軸・始まりの側)
     * @param negative 負の値 (軸から左へ伸びる)
     * @param rtl      右から左
     * @param axisColor 軸の色
     * @param hideText 値を表示しない
     */
    record Bar(double start, double end, double axis, Color fill, Color border, boolean gradient, boolean negative, boolean rtl,
               Color axisColor, boolean hideText) {}

    /** しきい値。 */
    private record Cfvo(String type, String val) {}

    /** 1 つの規則。 */
    private static final class Rule {
        List<CellRangeAddress> ranges = new ArrayList<>();
        Cfvo min = new Cfvo("min", null), max = new Cfvo("max", null);
        Color fill = new Color(0x638EC6), border, negFill, negBorder, axisColor = Color.BLACK;
        boolean showValue = true, gradient = true, rtl;
        double minLength = 0.1, maxLength = 0.9;
        String axisPosition = "none";
        int priority;
        /** 計算した最小・最大 (null は未計算)。 */
        double[] lohi;
    }

    private final List<Rule> rules = new ArrayList<>();
    private final XSSFSheet sheet;

    private DataBars(XSSFSheet sheet) { this.sheet = sheet; }

    /** シートのデータバーの規則を読む。なければ null。 */
    static DataBars parse(XSSFSheet sheet, List<CfExt.X14Rule> x14, Function<Object, Color> color) {
        DataBars out = new DataBars(sheet);
        Map<String, Rule> byId = new HashMap<>();
        for (CTConditionalFormatting cf : sheet.getCTWorksheet().getConditionalFormattingArray()) {
            List<CellRangeAddress> ranges = CfExt.ranges(CfExt.sqref(cf.getSqref()));
            for (CTCfRule r : cf.getCfRuleArray()) {
                if (r.getType() != STCfType.DATA_BAR || !r.isSetDataBar()) continue;
                CTDataBar db = r.getDataBar();
                Rule rule = new Rule();
                rule.ranges = ranges;
                CTCfvo[] vs = db.getCfvoArray();
                if (vs.length > 0) rule.min = new Cfvo(vs[0].getType().toString(), vs[0].getVal());
                if (vs.length > 1) rule.max = new Cfvo(vs[1].getType().toString(), vs[1].getVal());
                if (db.isSetShowValue()) rule.showValue = db.getShowValue();
                if (db.isSetMinLength()) rule.minLength = db.getMinLength() / 100.0;
                if (db.isSetMaxLength()) rule.maxLength = db.getMaxLength() / 100.0;
                Color c = db.getColor() == null ? null : color.apply(db.getColor());
                if (c != null) rule.fill = c;
                rule.priority = r.getPriority();
                out.rules.add(rule);
                for (String id : CfExt.x14Ids(r)) byId.put(id, rule);
            }
        }
        for (CfExt.X14Rule xr : x14) {
            Element r = xr.rule();
            if (!"dataBar".equals(r.getAttribute("type"))) continue;
            Element db = Xml.child(r, "dataBar");
            if (db == null) continue;
            // 通常の規則を拡張で補う (色・値の表示は通常の規則にある)
            Rule rule = byId.get(r.getAttribute("id"));
            if (rule == null) {
                rule = new Rule();
                rule.priority = Xml.intAttr(r, "priority", 0);
                out.rules.add(rule);
            }
            rule.ranges = xr.ranges();
            rule.minLength = Xml.intAttr(db, "minLength", 10) / 100.0;
            rule.maxLength = Xml.intAttr(db, "maxLength", 90) / 100.0;
            rule.gradient = !"0".equals(db.getAttribute("gradient")) && !"false".equals(db.getAttribute("gradient"));
            boolean border = "1".equals(db.getAttribute("border")) || "true".equals(db.getAttribute("border"));
            rule.rtl = "rightToLeft".equals(db.getAttribute("direction"));
            rule.axisPosition = db.hasAttribute("axisPosition") ? db.getAttribute("axisPosition") : "automatic";
            boolean negSame = "1".equals(db.getAttribute("negativeBarColorSameAsPositive"));
            boolean negBorderSame = !db.hasAttribute("negativeBarBorderColorSameAsPositive")
                    || "1".equals(db.getAttribute("negativeBarBorderColorSameAsPositive"));
            List<Element> vs = Xml.children(db, "cfvo");
            if (vs.size() > 0) rule.min = new Cfvo(vs.get(0).getAttribute("type"), Xml.text(Xml.child(vs.get(0), "f")));
            if (vs.size() > 1) rule.max = new Cfvo(vs.get(1).getAttribute("type"), Xml.text(Xml.child(vs.get(1), "f")));
            Color fc = color.apply(Xml.child(db, "fillColor"));
            if (fc != null) rule.fill = fc;
            Color bc = color.apply(Xml.child(db, "borderColor"));
            rule.border = border ? (bc != null ? bc : rule.fill) : null;
            Color nf = color.apply(Xml.child(db, "negativeFillColor"));
            rule.negFill = negSame ? rule.fill : nf != null ? nf : Color.RED;
            Color nb = color.apply(Xml.child(db, "negativeBorderColor"));
            rule.negBorder = !border ? null : negBorderSame ? rule.border : nb != null ? nb : Color.RED;
            Color ac = color.apply(Xml.child(db, "axisColor"));
            if (ac != null) rule.axisColor = ac;
        }
        if (out.rules.isEmpty()) return null;
        out.rules.sort((a, b) -> Integer.compare(a.priority, b.priority));
        return out;
    }

    /** セルの棒 (なければ null)。stopAt より優先順位の低い (priority の大きい) 規則は使わない (上の規則の「条件を満たす場合は停止」)。 */
    Bar at(int r, int c, BiFunction<Integer, Integer, Double> value, Function<String, Double> formula, int stopAt) {
        for (Rule rule : rules) {
            if (!CfExt.in(rule.ranges, r, c)) continue;
            if (rule.priority > stopAt) continue;
            Double v = value.apply(r, c);
            if (v == null) return null;
            if (rule.lohi == null) {
                double[] sorted = CfExt.values(sheet, rule.ranges, value);
                double lo = CfExt.threshold(rule.min.type(), rule.min.val(), sorted, formula);
                double hi = CfExt.threshold(rule.max.type(), rule.max.val(), sorted, formula);
                rule.lohi = new double[] {Double.isNaN(lo) ? 0 : lo, Double.isNaN(hi) ? 0 : hi};
            }
            return bar(rule, v);
        }
        return null;
    }

    /** 値から棒の位置を求める。 */
    private static Bar bar(Rule rule, double v) {
        double lo = rule.lohi[0], hi = rule.lohi[1];
        double axis = Double.NaN;
        if (rule.axisPosition.equals("middle")) axis = 0.5;
        else if (rule.axisPosition.equals("automatic") && lo < 0) axis = hi <= 0 ? 1 : -lo / (hi - lo);
        double start, end;
        boolean neg = false;
        if (Double.isNaN(axis)) {
            // 軸なし: 最小から最大を、最短〜最長の長さに対応させる
            double f = hi > lo ? (v - lo) / (hi - lo) : 1;
            f = Math.max(0, Math.min(1, f));
            start = 0;
            end = rule.minLength + f * (rule.maxLength - rule.minLength);
        } else {
            // 軸あり: 正と負を同じ縮尺で、軸から伸ばす
            // 自動の軸は 0 をはさんだ範囲 (すべて負なら 0 〜 最小。hi − lo では最長の棒がセルの外へ出る)
            double scale = rule.axisPosition.equals("middle") ? 2 * Math.max(Math.abs(lo), Math.abs(hi)) : Math.max(0, hi) - Math.min(0, lo);
            if (scale <= 0) scale = 1;
            if (v >= 0) { start = axis; end = axis + Math.min(v, Math.max(0, hi)) / scale; }
            else { neg = true; end = axis; start = axis - Math.min(-v, Math.max(0, -lo)) / scale; }
        }
        Color fill = neg ? rule.negFill : rule.fill;
        Color border = neg ? rule.negBorder : rule.border;
        return new Bar(start, end, axis, fill, border, rule.gradient, neg, rule.rtl, rule.axisColor, !rule.showValue);
    }
}
