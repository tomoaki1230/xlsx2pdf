package xlsx2pdf.render.formula;

import xlsx2pdf.render.Recalc;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/**
 * Excel 2024 日本語版で計算させたブック (src/test/resources/excel-grid) の値と、--recalc の結果を比べる。
 * ブックの Z 列に数式 (と Excel の計算結果)、AD 列以降に参照先のデータがある。結果を消してから再計算する。
 */
class ExcelGridTest {
    /** Excel の反復の打ち切りや内部の癖で、結果が揃わない式 (理由は設計書.md)。 */
    private static final Set<String> SKIP_FUNCTIONS = Set.of("XIRR", "VDB", "TYPE");

    /** 同じく揃わない式。Excel の t 分布の逆関数は、自由度が 1 億では 3e-9 ずれた値 (真の値より小さい) を返す。 */
    private static boolean skipFormula(String f) {
        return f.contains("100000000)") && f.contains("T.INV");
    }

    private static String show(Cell c) {
        if (c.getCellType() != CellType.FORMULA) return "?";
        return switch (c.getCachedFormulaResultType()) {
            case STRING -> "S:" + c.getStringCellValue();
            case NUMERIC -> "N:" + c.getNumericCellValue();
            case BOOLEAN -> "B:" + c.getBooleanCellValue();
            case ERROR -> "E:" + FormulaError.forInt(c.getErrorCellValue()).getString();
            default -> "?";
        };
    }

    private static double tolerance = 1e-9;

    private static boolean same(String got, String expected) {
        if (got.equals(expected)) return true;
        if (got.startsWith("N:") && expected.startsWith("N:")) {
            double x = Double.parseDouble(got.substring(2)), y = Double.parseDouble(expected.substring(2));
            return Math.abs(x - y) <= Math.max(tolerance, Math.max(Math.abs(x), Math.abs(y)) * tolerance);
        }
        return false;
    }

    private static List<String> compare(String resource) throws Exception {
        return compare(resource, 1e-9, SKIP_FUNCTIONS);
    }

    private static List<String> compare(String resource, double tol, Set<String> skip) throws Exception {
        tolerance = tol;
        List<String> bad = new ArrayList<>();
        try (InputStream in = ExcelGridTest.class.getResourceAsStream("/excel-grid/" + resource);
             XSSFWorkbook wb = new XSSFWorkbook(in)) {
            XSSFSheet s = wb.getSheet("Data");
            int last = s.getLastRowNum();
            List<String> expected = new ArrayList<>(), formulas = new ArrayList<>();
            for (int i = 0; i <= last; i++) {
                Row r = s.getRow(i);
                Cell c = r == null ? null : r.getCell(25);
                if (c == null || c.getCellType() != CellType.FORMULA) {
                    expected.add(null);
                    formulas.add(null);
                    continue;
                }
                String shown = show(c);
                // 試験の式に定義されていない名前を書いたもの (#NAME?) は比べない
                expected.add(shown.equals("E:#NAME?") ? null : shown);
                String f = c.getCellFormula();
                formulas.add(f);
                String name = f.replaceAll("^_xlfn\\.", "").replaceAll("\\(.*", "");
                if (expected.get(i) == null) continue;
                if (skip.contains(name) || skipFormula(f)) {
                    expected.set(i, null);
                    continue;
                }
                // 保存された結果を消して、計算できなかった式が「保存された結果のまま」で通らないようにする
                XSSFCell xc = (XSSFCell) c;
                xc.getCTCell().unsetV();
                if (xc.getCTCell().isSetT()) xc.getCTCell().unsetT();
            }
            Recalc.run(wb, null);
            for (int i = 0; i <= last; i++) {
                if (expected.get(i) == null) continue;
                String got = show(s.getRow(i).getCell(25));
                if (!same(got, expected.get(i))) bad.add(formulas.get(i) + " excel=" + expected.get(i) + " ours=" + got);
            }
        }
        return bad;
    }

    @Test
    void functions() throws Exception {
        assertEquals(List.of(), compare("functions.xlsx"));
    }

    @Test
    void romanAndIpmt() throws Exception {
        assertEquals(List.of(), compare("roman-ipmt.xlsx"));
    }

    @Test
    void aggregateOptions() throws Exception {
        assertEquals(List.of(), compare("aggregate.xlsx"));
    }

    @Test
    void distributionBoundaries() throws Exception {
        assertEquals(List.of(), compare("boundary.xlsx"));
    }

    @Test
    void convert() throws Exception {
        assertEquals(List.of(), compare("convert.xlsx"));
    }

    @Test
    void randomArguments() throws Exception {
        assertEquals(List.of(), compare("random.xlsx"));
    }

    @Test
    void subtotal() throws Exception {
        assertEquals(List.of(), compare("subtotal.xlsx"));
    }

    /**
     * XIRR は Excel と同じ手順 (区間の探し方・中点の計算・止める条件・推定値 0・-1 より下の利回り・#NUM! の条件) なので、最後の桁まで一致する。
     * 約 1 万 4 千式のうち、日付が昇順でなく利回りが -100% より下になる 104 式 (xirr-known.txt) だけ合わない。
     */
    @Test
    void xirr() throws Exception {
        Set<String> known;
        try (InputStream in = ExcelGridTest.class.getResourceAsStream("/excel-grid/xirr-known.txt")) {
            known = Set.copyOf(new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8).lines().filter(l -> !l.isBlank()).toList());
        }
        for (String f : List.of("xirr-positive.xlsx", "xirr-random.xlsx", "xirr-guess.xlsx", "xirr-zero.xlsx", "xirr-stop.xlsx",
                "xirr-small-first.xlsx", "xirr-large-guess.xlsx", "xirr-below-minus1.xlsx", "xirr-guess-minus1.xlsx",
                "xirr-random2.xlsx", "xirr-near-zero.xlsx", "xirr-two-roots.xlsx", "xirr-long-term.xlsx", "xirr-far-root.xlsx",
                "xirr-step-limit.xlsx", "xirr-random3.xlsx", "xirr-rounding.xlsx", "xirr-date-order.xlsx", "xirr-random4.xlsx",
                "xirr-below-minus1-shuffled.xlsx", "xirr-below-minus1-three.xlsx")) {
            List<String> bad = compare(f, 1e-15, Set.of());
            bad.removeIf(b -> known.contains(b.substring(0, b.indexOf(" excel="))));
            assertEquals(List.of(), bad, f);
        }
    }
}
