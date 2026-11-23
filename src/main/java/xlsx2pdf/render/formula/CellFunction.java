package xlsx2pdf.render.formula;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import org.apache.poi.ss.formula.eval.AreaEval;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.RefEval;
import org.apache.poi.ss.formula.eval.StringEval;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.formula.functions.Function;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import static xlsx2pdf.render.formula.FnSupport.*;
import static xlsx2pdf.render.formula.DateFunctions.*;

/** CELL (セルの情報)。CELL("format") の書式コードの判定を含む。 */
final class CellFunction {
    private CellFunction() {
    }


    static final class CellFn implements Function {
        @Override
        public ValueEval evaluate(ValueEval[] args, int srcRow, int srcCol) {
            try {
                if (args.length < 1 || args.length > 2) return ErrorEval.VALUE_INVALID;
                String info = OperandResolver.coerceValueToString(OperandResolver.getSingleValue(args[0], srcRow, srcCol)).toLowerCase(Locale.ROOT);
                int row = srcRow, col = srcCol, sheet = -1;
                ValueEval inner = null;
                if (args.length == 1) {
                    // 参照がないとき Excel は「最後に変更したセル」を返す。再計算では、ブックを開いたときの選択中のセルになる
                    XSSFWorkbook w0 = FormulaContext.workbook();
                    if (w0 != null) {
                        int ai = w0.getActiveSheetIndex();
                        if (ai >= 0 && ai < w0.getNumberOfSheets()) {
                            org.apache.poi.ss.util.CellAddress ca = w0.getSheetAt(ai).getActiveCell();
                            sheet = ai;
                            if (ca != null) {
                                row = ca.getRow();
                                col = ca.getColumn();
                            } else {
                                row = 0;
                                col = 0;
                            }
                        }
                    }
                }
                if (args.length == 2) {
                    ValueEval r = args[1];
                    if (r instanceof RefEval re) {
                        row = re.getRow();
                        col = re.getColumn();
                        sheet = re.getFirstSheetIndex();
                        inner = re.getInnerValueEval(sheet);
                    } else if (r instanceof AreaEval ar) {
                        row = ar.getFirstRow();
                        col = ar.getFirstColumn();
                        inner = ar.getRelativeValue(0, 0);
                    } else {
                        return ErrorEval.VALUE_INVALID;
                    }
                }
                XSSFWorkbook wb = FormulaContext.workbook();
                XSSFSheet sh = null;
                if (wb != null && sheet >= 0 && sheet < wb.getNumberOfSheets()) sh = wb.getSheetAt(sheet);
                XSSFCell cell = null;
                if (sh != null && sh.getRow(row) != null) cell = sh.getRow(row).getCell(col);
                switch (info) {
                    case "address": {
                        String ad = "$" + org.apache.poi.ss.util.CellReference.convertNumToColString(col) + "$" + (row + 1);
                        if (sh != null && sheet != FormulaContext.currentSheet() && FormulaContext.currentSheet() >= 0) {
                            String book = FormulaContext.filePath() == null ? "Book1" : new java.io.File(FormulaContext.filePath()).getName();
                            String nm = "[" + book + "]" + sh.getSheetName();
                            boolean plain = sh.getSheetName().matches("[A-Za-z_][A-Za-z0-9_.]*");
                            return new StringEval((plain ? nm : "'" + nm.replace("'", "''") + "'") + "!" + ad);
                        }
                        return new StringEval(ad);
                    }
                    case "col":
                        return new NumberEval(col + 1);
                    case "row":
                        return new NumberEval(row + 1);
                    case "contents":
                        return inner == null ? new NumberEval(0) : inner;
                    case "type":
                        if (inner == null || inner instanceof BlankEval) return new StringEval("b");
                        return new StringEval(inner instanceof StringEval ? "l" : "v");
                    case "filename": {
                        if (sh == null) sh = FormulaContext.currentSheet() >= 0 && wb != null ? wb.getSheetAt(FormulaContext.currentSheet()) : null;
                        if (sh == null || FormulaContext.filePath() == null) return new StringEval("");
                        java.io.File f = new java.io.File(FormulaContext.filePath());
                        String dir = f.getParent() == null ? "" : f.getParent();
                        java.util.regex.Matcher wm = java.util.regex.Pattern.compile("/mnt/([a-z])(/.*)?").matcher(dir);
                        // WSL のパスは Windows の形 (C:\dir\[file]sheet) にする
                        if (wm.matches()) dir = wm.group(1).toUpperCase(Locale.ROOT) + ":" + (wm.group(2) == null ? "\\" : wm.group(2).replace('/', '\\'));
                        String sep = dir.contains("\\") || dir.matches("[A-Za-z]:") ? "\\" : java.io.File.separator;
                        return new StringEval((dir.isEmpty() ? "" : dir + sep) + "[" + f.getName() + "]" + sh.getSheetName());
                    }
                    case "protect":
                        return new NumberEval(cell != null && cell.getCellStyle().getLocked() ? 1 : (cell == null ? 1 : 0));
                    case "prefix": {
                        if (!(inner instanceof StringEval) || cell == null) return new StringEval("");
                        switch (cell.getCellStyle().getAlignment()) {
                            case GENERAL:
                            case LEFT: return new StringEval("'");
                            case RIGHT: return new StringEval("\"");
                            case CENTER: return new StringEval("^");
                            case FILL: return new StringEval("\\");
                            default: return new StringEval("");
                        }
                    }
                    case "width": {
                        if (sh == null) return new NumberEval(8);
                        int w = sh.getColumnWidth(col);
                        return new NumberEval(Math.floor(w / 256.0));
                    }
                    case "color":
                    case "parentheses": {
                        String f = cell == null ? "General" : cell.getCellStyle().getDataFormatString();
                        return new NumberEval((info.equals("color") ? isRed(f) : hasParen(f)) ? 1 : 0);
                    }
                    case "format":
                        if (cell == null) return new StringEval("G");
                        // 日本語版 Excel の組み込み書式 (書式コードがファイルにない 30〜36・50〜58 番)
                        String fs = switch (cell.getCellStyle().getDataFormat()) {
                            case 30, 31 -> "yyyy/m/d";
                            case 32 -> "h:mm";
                            case 33 -> "h:mm:ss";
                            case 36, 50, 51, 52, 53, 54, 55, 56, 57 -> "ggge\"年\"m\"月\"d\"日\"";
                            case 58 -> "m\"月\"d\"日\"";
                            default -> cell.getCellStyle().getDataFormatString();
                        };
                        return new StringEval(formatCode(fs));
                    default:
                        return ErrorEval.VALUE_INVALID;
                }
            } catch (EvaluationException e) {
                return e.getErrorEval();
            }
        }
    }

    static String[] sections(String f) {
        // 引用符の中の ; は区切りでない
        List<String> out = new ArrayList<>();
        StringBuilder sb = new StringBuilder();
        boolean q = false;
        for (int i = 0; i < f.length(); i++) {
            char c = f.charAt(i);
            if (c == '"') q = !q;
            if (c == '\\' || c == '_' || c == '*') {
                sb.append(c);
                if (i + 1 < f.length()) sb.append(f.charAt(++i));
                continue;
            }
            if (c == ';' && !q) {
                out.add(sb.toString());
                sb.setLength(0);
            } else {
                sb.append(c);
            }
        }
        out.add(sb.toString());
        return out.toArray(new String[0]);
    }

    /** 2 つ以上の区間があり、正または負の区間に [Red] がある。 */
    static boolean isRed(String f) {
        String[] p = sections(f);
        if (p.length < 2) return false;
        return p[0].toLowerCase(Locale.ROOT).contains("[red]") || p[1].toLowerCase(Locale.ROOT).contains("[red]");
    }

    /** 正の区間に (がある ("_(" と引用符の中は数えない)。 */
    static boolean hasParen(String f) {
        String p = sections(f)[0].replaceAll("\"[^\"]*\"", "").replaceAll("[_\\\\*].", "");
        return p.contains("(");
    }

    /** CELL("format") の書式コード (日本語版 Excel 2024 の結果に合わせた)。 */
    static String formatCode(String f) {
        String sec = sections(f)[0];
        String suffix = (isRed(f) ? "-" : "") + (hasParen(f) ? "()" : "");
        // 引用符・\x・_x・*x・[...] を除く (経過時間の [h] は先に調べる)
        String raw = sec.replaceAll("\"[^\"]*\"", "").replaceAll("[_\\\\*].", "");
        if (raw.matches("(?i).*\\[(h+|m+|s+)\\].*")) return "G" + suffix;
        String nb = sec.replaceAll("\\[[^\\]]*\\]", "");
        boolean yen = nb.contains("¥") || nb.contains("￥") || nb.contains("\\#");
        String p = raw.replaceAll("\\[[^\\]]*\\]", "");
        String low = p.toLowerCase(Locale.ROOT).replaceAll("(?<=s)\\.0+", "");  // 秒の小数 (s.0) は桁でない
        if (sec.toLowerCase(Locale.ROOT).contains("[$-f4")) return "G";
        if (low.isEmpty() || low.equals("general") || low.matches("@+")) return "G";
        if (low.contains("?/")) return "G";
        boolean digits = low.matches(".*[0#].*");
        if (!digits || low.matches(".*[yd].*|.*h.*|.*s.*|.*g.*") && !low.matches(".*e[+-].*")) {
            String c = dateCode(low);
            if (c != null) return c + suffix;
            if (!digits) return "G" + suffix;
        }
        int dec = 0;
        int dot = p.indexOf('.');
        if (dot >= 0) {
            for (int i = dot + 1; i < p.length() && (p.charAt(i) == '0' || p.charAt(i) == '#' || p.charAt(i) == '?'); i++) dec++;
        }
        if (low.matches(".*e[+-].*")) return "S" + dec + suffix;
        if (p.contains("%")) return "P" + dec + suffix;
        if (yen) return "C" + dec + suffix;
        if (low.matches(".*[#0?],[#0?].*")) return "," + dec + suffix;
        return "F" + dec + suffix;
    }

    /** 日付・時刻の書式の D1〜D9 (該当しなければ null)。 */
    static String dateCode(String low) {
        boolean ampm = low.contains("am/pm") || low.contains("a/p");
        String t = low.replace("am/pm", "").replace("a/p", "");
        boolean y = t.contains("y"), era = t.contains("g") || t.matches(".*e.*");
        boolean day = t.matches(".*(?<!d)dd?(?!d).*"), h = t.contains("h"), s = t.contains("s");
        // m: 直前が h・直後が s なら分、それ以外は月
        boolean month = false, minute = false;
        for (int i = 0; i < t.length(); i++) {
            if (t.charAt(i) != 'm') continue;
            int j = i;
            while (j < t.length() && t.charAt(j) == 'm') j++;
            String before = t.substring(0, i).replaceAll("[^a-z]", ""), after = t.substring(j).replaceAll("[^a-z]", "");
            if (before.endsWith("h") || after.startsWith("s")) minute = true;
            else month = true;
            i = j - 1;
        }
        if (era && (month || day || y)) return "D4";
        if (y && month && day) return "D1";
        if (y && month) return "D2";
        if (month && day) return "D3";
        if (h && (minute || s)) return s ? (ampm ? "D6" : "D8") : (ampm ? "D7" : "D9");
        return null;
    }
}
