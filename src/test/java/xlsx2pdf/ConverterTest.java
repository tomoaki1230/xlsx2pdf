package xlsx2pdf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ConverterTest {
    /** シートごとに A1 に文字を入れたブック。 */
    private static byte[] book(String... sheets) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            for (String name : sheets) {
                XSSFSheet s = wb.createSheet(name);
                s.createRow(0).createCell(0).setCellValue("text-" + name);
            }
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] convert(byte[] xlsx, Converter.Options opt) throws Exception {
        ByteArrayOutputStream pdf = new ByteArrayOutputStream();
        new Converter(opt).convert(new ByteArrayInputStream(xlsx), pdf);
        return pdf.toByteArray();
    }

    private static String text(byte[] pdf) throws IOException {
        try (PDDocument d = Loader.loadPDF(pdf)) {
            return new PDFTextStripper().getText(d);
        }
    }

    private static int pages(byte[] pdf) throws IOException {
        try (PDDocument d = Loader.loadPDF(pdf)) {
            return d.getNumberOfPages();
        }
    }

    @Test
    void convertsAllVisibleSheets() throws Exception {
        byte[] pdf = convert(book("a", "b"), new Converter.Options());
        assertEquals(2, pages(pdf));
        String t = text(pdf);
        assertTrue(t.contains("text-a") && t.contains("text-b"), t);
    }

    @Test
    void convertsOnlyNamedSheet() throws Exception {
        Converter.Options opt = new Converter.Options();
        opt.sheets.add("b");
        byte[] pdf = convert(book("a", "b"), opt);
        assertEquals(1, pages(pdf));
        String t = text(pdf);
        assertTrue(t.contains("text-b"), t);
        assertFalse(t.contains("text-a"), t);
    }

    @Test
    void writesFileFromFile(@TempDir Path dir) throws Exception {
        File in = dir.resolve("in.xlsx").toFile(), out = dir.resolve("out.pdf").toFile();
        Files.write(in.toPath(), book("a"));
        new Converter(new Converter.Options()).convert(in, out);
        assertEquals(1, pages(Files.readAllBytes(out.toPath())));
    }

    @Test
    void writesNothingWhenInputIsNotXlsx() {
        ByteArrayOutputStream pdf = new ByteArrayOutputStream();
        assertThrows(Exception.class, () -> new Converter(new Converter.Options())
                .convert(new ByteArrayInputStream("not xlsx".getBytes()), pdf));
        assertEquals(0, pdf.size());
    }

    @Test
    void stopsAtPageLimit() {
        Converter.Options opt = new Converter.Options();
        opt.maxPages = 1;
        assertThrows(LimitExceededException.class, () -> convert(book("a", "b", "c"), opt));
    }

    @Test
    void stopsAtInputLimit() {
        Converter.Options opt = new Converter.Options();
        opt.maxInputBytes = 10;
        assertThrows(LimitExceededException.class, () -> convert(book("a"), opt));
    }

    @Test
    void roundsToFifteenDigitsBeforeFormatting() throws Exception {
        byte[] xlsx;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet s = wb.createSheet("a");
            var st = wb.createCellStyle();
            st.setDataFormat(wb.createDataFormat().getFormat("0.00"));
            var c = s.createRow(0).createCell(0);
            c.setCellValue(1.4449999999999998);
            c.setCellStyle(st);
            wb.write(out);
            xlsx = out.toByteArray();
        }
        String t = text(convert(xlsx, new Converter.Options()));
        assertTrue(t.contains("1.45"), t);
    }

    @Test
    void drawsMergedCellWhoseFirstRowIsHidden() throws Exception {
        byte[] xlsx;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet s = wb.createSheet("a");
            s.createRow(1).createCell(1).setCellValue("merged-text");
            s.getRow(1).setZeroHeight(true);
            s.createRow(2);
            s.createRow(3);
            s.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(1, 3, 1, 3));
            wb.write(out);
            xlsx = out.toByteArray();
        }
        String t = text(convert(xlsx, new Converter.Options()));
        assertTrue(t.contains("merged-text"), t);
    }

    @Test
    void returnsNoWarningsForNormalBook() throws Exception {
        Converter conv = new Converter(new Converter.Options());
        conv.convert(new ByteArrayInputStream(book("a")), new ByteArrayOutputStream());
        assertEquals(java.util.List.of(), conv.getWarnings());
    }

    @Test
    void forbidsEditing() throws Exception {
        Converter.Options opt = new Converter.Options();
        opt.noEdit = true;
        opt.ownerPassword = "owner-secret";
        byte[] pdf = convert(book("a"), opt);
        try (PDDocument d = Loader.loadPDF(pdf)) {
            assertTrue(d.isEncrypted());
            var ap = d.getCurrentAccessPermission();
            assertFalse(ap.canModify());
            assertFalse(ap.canModifyAnnotations());
            assertTrue(ap.canPrint());
            assertTrue(ap.canExtractContent());
            assertTrue(new PDFTextStripper().getText(d).contains("text-a"));
        }
        try (PDDocument d = Loader.loadPDF(pdf, "owner-secret")) {
            assertTrue(d.getCurrentAccessPermission().isOwnerPermission());
        }
    }

    @Test
    void forbidsEditingOfFileOutput(@TempDir Path dir) throws Exception {
        File in = dir.resolve("in.xlsx").toFile(), out = dir.resolve("out.pdf").toFile();
        Files.write(in.toPath(), book("a"));
        Converter.Options opt = new Converter.Options();
        opt.noEdit = true;
        new Converter(opt).convert(in, out);
        try (PDDocument d = Loader.loadPDF(out)) {
            assertTrue(d.isEncrypted());
            assertFalse(d.getCurrentAccessPermission().canModify());
        }
    }

    @Test
    void parsesOptions() {
        java.util.List<String> files = new java.util.ArrayList<>();
        Converter.Options opt = Main.parseOptions(new String[] {"in.xlsx", "out.pdf", "--author-env", "--screen-dpi", "150%", "--sheet", "s"}, files);
        assertEquals(java.util.List.of("in.xlsx", "out.pdf"), files);
        assertTrue(opt.authorEnv);
        assertEquals(144, opt.screenDpi);
        assertEquals(java.util.List.of("s"), opt.sheets);
    }

    /** 保存された計算結果が古いブック: 既定は保存された値、--recalc は計算し直した値。計算できない数式は保存された値のまま。 */
    @Test
    void recalcReplacesStaleResults() throws Exception {
        byte[] xlsx;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet s = wb.createSheet("s");
            s.createRow(0).createCell(0).setCellValue(2);
            s.createRow(1).createCell(0).setCellValue(3);
            var f = s.createRow(2).createCell(0);
            f.setCellFormula("A1+A2");
            f.getCTCell().setV("999");          // 古い計算結果
            var g = s.createRow(3).createCell(0);
            g.setCellFormula("\"v\"&A1");
            g.getCTCell().setT(org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType.STR);
            g.getCTCell().setV("stale");
            wb.write(out);
            xlsx = out.toByteArray();
        }
        String before = text(convert(xlsx, new Converter.Options()));
        assertTrue(before.contains("999") && before.contains("stale"), before);
        Converter.Options opt = Main.parseOptions(new String[] {"--recalc"}, new java.util.ArrayList<>());
        assertTrue(opt.recalc);
        String after = text(convert(xlsx, opt));
        assertTrue(after.contains("5") && after.contains("v2"), after);
        assertFalse(after.contains("999") || after.contains("stale"), after);
    }

    /** sheet1.xml の文字列を置き換えたブック。 */
    private static byte[] replaceInSheet(byte[] xlsx, String from, String to) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var zin = new java.util.zip.ZipInputStream(new ByteArrayInputStream(xlsx));
             var zout = new java.util.zip.ZipOutputStream(out)) {
            for (java.util.zip.ZipEntry e; (e = zin.getNextEntry()) != null; ) {
                byte[] b = zin.readAllBytes();
                if (e.getName().equals("xl/worksheets/sheet1.xml"))
                    b = new String(b, java.nio.charset.StandardCharsets.UTF_8).replace(from, to).getBytes(java.nio.charset.StandardCharsets.UTF_8);
                zout.putNextEntry(new java.util.zip.ZipEntry(e.getName()));
                zout.write(b);
                zout.closeEntry();
            }
        }
        return out.toByteArray();
    }

    /** --recalc で POI にない関数は、保存された結果のまま印刷する。 */
    @Test
    void recalcKeepsUnknownFunctions() throws Exception {
        byte[] xlsx;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            XSSFSheet s = wb.createSheet("s");
            var f = s.createRow(0).createCell(0);
            f.setCellFormula("1+1");
            f.getCTCell().setV("4242");
            wb.write(out);
            xlsx = out.toByteArray();
        }
        xlsx = replaceInSheet(xlsx, "<f>1+1</f>", "<f>FOOBARX(1)</f>");
        Converter.Options opt = new Converter.Options();
        opt.recalc = true;
        assertTrue(text(convert(xlsx, opt)).contains("4242"));
    }
}
