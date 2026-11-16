package xlsx2pdf.render;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class ExcelFormatTest {
    @Test
    void dbnumRewritesDigitRuns() {
        assertEquals("千二百三十四", ExcelFormat.kanjiInt(1234, 1));
        assertEquals("壱阡弐百参拾四", ExcelFormat.kanjiInt(1234, 2));
        assertEquals("一万十", ExcelFormat.kanjiInt(10010, 1));
        assertEquals("明治三十三年一月〇日", ExcelFormat.dbnum("明治33年1月0日", 1));
        assertEquals("-１３", ExcelFormat.dbnum("-13", 3));
        assertEquals("十二.五", ExcelFormat.dbnum("12.5", 1));
    }

    @Test
    void generalFitsByDigitCount() {
        assertEquals("3.1416", ValueFormatter.generalFit(3.14159265358979, 6));
        assertEquals("1E+08", ValueFormatter.generalFit(123456789.123, 6));
        assertEquals("-1E-03", ValueFormatter.generalFit(-0.000987654, 6));
        assertEquals("-0.0009877", ValueFormatter.generalFit(-0.000987654, 10));
        assertEquals("1.2345E-05", ValueFormatter.generalFit(0.000012345, 10));
        assertEquals("0", ValueFormatter.generalFit(0.000012345, 3));
        assertEquals("-0", ValueFormatter.generalFit(-0.000987654, 3));
    }
}
