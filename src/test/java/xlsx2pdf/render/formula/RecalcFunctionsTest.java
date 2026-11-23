package xlsx2pdf.render.formula;

import xlsx2pdf.render.Recalc;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class RecalcFunctionsTest {
    /** 数式を A1 に入れて再計算し、結果を文字列で返す。 */
    private static String eval(String formula) throws Exception {
        return eval(formula, false);
    }

    private static String eval(String formula, boolean d1904) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            if (d1904) (wb.getCTWorkbook().isSetWorkbookPr() ? wb.getCTWorkbook().getWorkbookPr() : wb.getCTWorkbook().addNewWorkbookPr()).setDate1904(true);
            XSSFSheet s = wb.createSheet("Sheet1");
            Row r = s.createRow(0);
            r.createCell(1).setCellValue("a-b-c");
            r.createCell(2).setCellValue(45000); // 2023-03-15
            r.createCell(3).setCellValue(45500);
            r.createCell(4).setCellValue(true);
            XSSFCell c = (XSSFCell) r.createCell(0);
            c.setCellFormula(formula);
            Recalc.run(wb, null);
            if (c.getCachedFormulaResultType() == CellType.ERROR) return "ERR" + c.getErrorCellValue();
            if (c.getCachedFormulaResultType() == CellType.STRING) return c.getStringCellValue();
            if (c.getCachedFormulaResultType() == CellType.BOOLEAN) return String.valueOf(c.getBooleanCellValue());
            return String.valueOf((long) c.getNumericCellValue());
        }
    }

    @Test
    void textBeforeAfter() throws Exception {
        assertEquals("a", eval("TEXTBEFORE(B1,\"-\")"));
        assertEquals("a-b", eval("TEXTBEFORE(B1,\"-\",2)"));
        assertEquals("a-b", eval("TEXTBEFORE(B1,\"-\",-1)"));
        assertEquals("b-c", eval("TEXTAFTER(B1,\"-\")"));
        assertEquals("c", eval("TEXTAFTER(B1,\"-\",-1)"));
        assertEquals("x", eval("TEXTAFTER(B1,\"?\",1,0,0,\"x\")"));
        assertEquals("ERR42", eval("TEXTAFTER(B1,\"?\")"));
        assertEquals("a-b-c", eval("TEXTBEFORE(B1,\"-\",3,0,1)"));
    }

    @Test
    void datedifXorIsoweek() throws Exception {
        assertEquals("1", eval("DATEDIF(C1,D1,\"Y\")"));
        assertEquals("16", eval("DATEDIF(C1,D1,\"M\")"));
        assertEquals("500", eval("DATEDIF(C1,D1,\"D\")"));
        assertEquals("ERR36", eval("DATEDIF(D1,C1,\"D\")"));
        assertEquals("true", eval("XOR(E1,FALSE)"));
        assertEquals("false", eval("XOR(TRUE,TRUE)"));
        assertEquals("11", eval("ISOWEEKNUM(C1)"));
    }

    @Test
    void cell() throws Exception {
        assertEquals("$B$1", eval("CELL(\"address\",B1)"));
        assertEquals("2", eval("CELL(\"col\",B1)"));
        assertEquals("l", eval("CELL(\"type\",B1)"));
    }

    /** 日本語版 Excel 2024 の CELL("format")・"color"・"parentheses" の実測 (src/test/resources/cell-format-excel.tsv) と一致する。 */
    @Test
    void cellFormatMatchesExcel() throws Exception {
        java.util.List<String> bad = new java.util.ArrayList<>();
        try (var in = getClass().getResourceAsStream("/cell-format-excel.tsv");
             var r = new java.io.BufferedReader(new java.io.InputStreamReader(in, java.nio.charset.StandardCharsets.UTF_8))) {
            for (String line; (line = r.readLine()) != null; ) {
                String[] f = line.split("\t", -1);
                String code = CellFunction.formatCode(f[0]);
                String color = CellFunction.isRed(f[0]) ? "1" : "0", paren = CellFunction.hasParen(f[0]) ? "1" : "0";
                if (!code.equals(f[1]) || !color.equals(f[2]) || !paren.equals(f[3]))
                    bad.add(f[0] + " excel=" + f[1] + "/" + f[2] + "/" + f[3] + " ours=" + code + "/" + color + "/" + paren);
            }
        }
        assertEquals(java.util.List.of(), bad);
    }

    @Test
    void datedifYdAndEmptyDelimiter() throws Exception {
        // Excel の実測: 開始日の年のうるう日を、月の端数まで進んだところで数える
        assertEquals("61", eval("DATEDIF(DATE(2020,1,28),DATE(2021,3,29),\"YD\")"));
        assertEquals("33", eval("DATEDIF(DATE(2020,1,28),DATE(2021,3,2),\"YD\")"));
        assertEquals("68", eval("DATEDIF(DATE(2023,12,29),DATE(2026,3,7),\"YD\")"));
        assertEquals("360", eval("DATEDIF(DATE(2021,3,30),DATE(2024,3,24),\"YD\")"));
        assertEquals("99", eval("DATEDIF(1,100,\"YD\")"));
        assertEquals("b", eval("TEXTBEFORE(\"b\",\"\")"));
        assertEquals("", eval("TEXTBEFORE(\"ab\",\"\")"));
        assertEquals("ab", eval("TEXTAFTER(\"ab\",\"\",5)"));
        assertEquals("$A$1", eval("CELL(\"address\")"));
    }

    /** 1904 年の日付システムのブックでは、POI の日付の関数に 1462 日のずれを補う (Excel の実測)。 */
    @Test
    void date1904() throws Exception {
        assertEquals("42369", eval("DATE(2020,1,1)", true));
        assertEquals("2024", eval("YEAR(43831)", true));
        assertEquals("3", eval("WEEKDAY(43831)", true));
        assertEquals("0", eval("DATE(1904,1,1)", true));
        assertEquals("ERR36", eval("DATE(1903,12,31)", true));
        assertEquals("43831", eval("DATE(2020,1,1)", false));
        assertEquals("400", eval("DAYS(0+400,0)", true));
        assertEquals("1904/1/2", eval("TEXT(1,\"yyyy/m/d\")", true));
    }

    /** VDB: 開始が整数ならすべて Excel と一致。開始が小数で定額法へ切り替えるときは、開始時点の簿価から数え直す (Excel で確認)。 */
    @Test
    void vdb() throws Exception {
        assertEquals("2100", eval("ROUND(VDB(2400,300,10,0,10),0)"));
        assertEquals("114", eval("ROUND(VDB(2400,300,10,0.25,0.5),4)"));
        assertEquals("547", eval("ROUND(VDB(2400,300,10,0.25,1.5),0)"));
        assertEquals("120", eval("ROUND(VDB(2400,300,10,0.25,0.5,2,TRUE),4)"));
    }
}
