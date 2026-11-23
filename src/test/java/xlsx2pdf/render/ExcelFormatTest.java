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
        // 数値の書式では 1 字ずつ、日付の書式では 4 桁の年と 0 始まりは 1 字ずつ (Excel の実測)
        assertEquals("一二三四", ExcelFormat.dbnum("1234", 1, 0));
        assertEquals("二〇二〇/〇一/十二", ExcelFormat.dbnum("2020/01/12", 1));
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

    /** 条件のある書式の負の符号と、無限大の # (Excel 2024 で確かめた値)。 */
    @Test
    void conditionalSectionsDropMinusOnlyForNegativeRanges() {
        ExcelFormat f = new ExcelFormat(false);
        assertEquals("5", f.format(-5, "[>0]0;0").text);
        assertEquals("5", f.format(-5, "[>-1]0;0").text);
        assertEquals("-5", f.format(-5, "[>5]0;0").text);
        assertEquals("-5", f.format(-5, "[<5]0;0").text);
        assertEquals("-5", f.format(-5, "[<=0]0;0").text);
        assertEquals("5", f.format(-5, "[<0]0;0").text);
        assertEquals("5", f.format(-5, "[<-1]0;0").text);
        assertEquals("-20", f.format(-20, "[=-5]0;0").text);
        assertEquals("-5", f.format(-5, "[>0]0;[<-9]0;0").text);
        assertEquals("########", f.format(Double.POSITIVE_INFINITY, "0.00").text);
    }

    /** 経過時間 (Excel 2024 で確かめた値):の書式は、ほかの要素が付くと 9999/12/31 を超える値を # で埋める。 */
    @Test
    void elapsedFormatsFillHashBeyondMaxDateOnlyWithOtherParts() {
        ExcelFormat f = new ExcelFormat(false);
        assertEquals("71003183:45", f.format(2958465.99, "[h]:mm").text);
        assertEquals("########", f.format(2958466, "[h]:mm").text);
        assertEquals("########", f.format(3000000, "[mm]:ss").text);
        assertEquals("########", f.format(3000000, "[h]:mm m").text);
        assertEquals("72000000", f.format(3000000, "[h]").text);
        assertEquals("4320000000", f.format(3000000, "[m]").text);
    }
}
