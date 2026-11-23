package xlsx2pdf.render;

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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconSet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.w3c.dom.Element;

/**
 * 条件付き書式のアイコンセット。通常の規則 (conditionalFormatting/cfRule) と、Excel 2010 以降の拡張
 * (3Stars、3Triangles、5Boxes と独自の組み合わせ) を読む。
 */
final class IconSets {
    /**
     * セルに描くアイコン。
     *
     * @param set      アイコンセットの名前 (3Arrows など)
     * @param index    セットの中の番号 (0 が最も低い値のアイコン)
     * @param iconOnly 値を表示しない
     */
    record Icon(String set, int index, boolean iconOnly) {}

    /** しきい値 (cfvo)。 */
    private record Cfvo(String type, String val, boolean gte) {}

    /** 1 つの規則。 */
    private static final class Rule {
        List<CellRangeAddress> ranges = new ArrayList<>();
        String set = "3TrafficLights1";
        boolean showValue = true, reverse;
        List<Cfvo> cfvos = new ArrayList<>();
        /** 独自の組み合わせ (x14:cfIcon) のセットと番号。null は既定。 */
        String[] customSet;
        int[] customId;
        int priority;
        /** 計算したしきい値。 */
        double[] thresholds;
    }

    private final List<Rule> rules = new ArrayList<>();
    private final XSSFSheet sheet;

    private IconSets(XSSFSheet sheet) { this.sheet = sheet; }

    /** シートのアイコンセットの規則を読む。なければ null。 */
    static IconSets parse(XSSFSheet sheet, List<CfExt.X14Rule> x14) {
        IconSets out = new IconSets(sheet);
        Map<String, Rule> byId = new HashMap<>();
        for (CTConditionalFormatting cf : sheet.getCTWorksheet().getConditionalFormattingArray()) {
            List<CellRangeAddress> ranges = CfExt.ranges(CfExt.sqref(cf.getSqref()));
            for (CTCfRule r : cf.getCfRuleArray()) {
                if (r.getType() != STCfType.ICON_SET || !r.isSetIconSet()) continue;
                CTIconSet is = r.getIconSet();
                Rule rule = new Rule();
                rule.ranges = ranges;
                if (is.isSetIconSet()) rule.set = is.getIconSet().toString();
                if (is.isSetShowValue()) rule.showValue = is.getShowValue();
                if (is.isSetReverse()) rule.reverse = is.getReverse();
                for (CTCfvo v : is.getCfvoArray()) rule.cfvos.add(new Cfvo(v.getType().toString(), v.getVal(), !v.isSetGte() || v.getGte()));
                rule.priority = r.getPriority();
                out.rules.add(rule);
                // 拡張に同じ規則の新しい形があれば、そちらで置き換える
                for (String id : CfExt.x14Ids(r)) byId.put(id, rule);
            }
        }
        for (CfExt.X14Rule xr : x14) {
            Element r = xr.rule();
            if (!"iconSet".equals(r.getAttribute("type"))) continue;
            Element is = Xml.child(r, "iconSet");
            if (is == null) continue;
            Rule rule = new Rule();
            rule.ranges = xr.ranges();
            if (is.hasAttribute("iconSet")) rule.set = is.getAttribute("iconSet");
            rule.showValue = !"0".equals(is.getAttribute("showValue")) && !"false".equals(is.getAttribute("showValue"));
            rule.reverse = "1".equals(is.getAttribute("reverse")) || "true".equals(is.getAttribute("reverse"));
            for (Element v : Xml.children(is, "cfvo")) {
                String gte = v.getAttribute("gte");
                rule.cfvos.add(new Cfvo(v.getAttribute("type"), Xml.text(Xml.child(v, "f")), !"0".equals(gte) && !"false".equals(gte)));
            }
            List<Element> icons = Xml.children(is, "cfIcon");
            if (!icons.isEmpty()) {
                rule.customSet = new String[icons.size()];
                rule.customId = new int[icons.size()];
                for (int i = 0; i < icons.size(); i++) {
                    rule.customSet[i] = icons.get(i).getAttribute("iconSet");
                    rule.customId[i] = Xml.intAttr(icons.get(i), "iconId", 0);
                }
            }
            rule.priority = Xml.intAttr(r, "priority", 0);
            Rule old = byId.get(r.getAttribute("id"));
            if (old != null) out.rules.remove(old);
            out.rules.add(rule);
        }
        if (out.rules.isEmpty()) return null;
        out.rules.sort((a, b) -> Integer.compare(a.priority, b.priority));
        return out;
    }

    /**
     * セルのアイコン。
     *
     * @param value   セルの数値 (数値でなければ null)
     * @param formula しきい値の数式を評価する。評価できなければ null
     * @param stopAt  これより優先順位の低い (priority の大きい) 規則は使わない (上の規則の「条件を満たす場合は停止」)
     * @return 描くアイコン。なければ null
     */
    Icon at(int r, int c, BiFunction<Integer, Integer, Double> value, Function<String, Double> formula, int stopAt) {
        for (Rule rule : rules) {
            if (!CfExt.in(rule.ranges, r, c)) continue;
            if (rule.priority > stopAt) continue;
            Double v = value.apply(r, c);
            if (v == null) return null;
            if (rule.thresholds == null) {
                double[] sorted = CfExt.values(sheet, rule.ranges, value);
                rule.thresholds = new double[rule.cfvos.size()];
                for (int i = 0; i < rule.cfvos.size(); i++) {
                    Cfvo cv = rule.cfvos.get(i);
                    rule.thresholds[i] = CfExt.threshold(cv.type(), cv.val(), sorted, formula);
                }
            }
            int n = rule.cfvos.size();
            int level = 0;
            for (int i = 1; i < n; i++) {
                double t = rule.thresholds[i];
                if (Double.isNaN(t)) continue;
                if (rule.cfvos.get(i).gte() ? v >= t : v > t) level = i;
            }
            if (rule.reverse) level = n - 1 - level;
            if (rule.customSet != null && level < rule.customSet.length) {
                String set = rule.customSet[level];
                if ("NoIcons".equals(set)) return null;
                return new Icon(set, rule.customId[level], !rule.showValue);
            }
            return new Icon(rule.set, level, !rule.showValue);
        }
        return null;
    }
}
