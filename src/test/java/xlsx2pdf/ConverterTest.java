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
    void parsesOptions() {
        java.util.List<String> files = new java.util.ArrayList<>();
        Converter.Options opt = Main.parseOptions(new String[] {"in.xlsx", "out.pdf", "--author-env", "--screen-dpi", "150%", "--sheet", "s"}, files);
        assertEquals(java.util.List.of("in.xlsx", "out.pdf"), files);
        assertTrue(opt.authorEnv);
        assertEquals(144, opt.screenDpi);
        assertEquals(java.util.List.of("s"), opt.sheets);
    }
}
