package xlsx2pdf.render;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class PrintSetupTest {
    @Test
    void usesFileDpiOnlyWhenRequested() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet s = wb.createSheet("a");
            s.getCTWorksheet().addNewPageSetup().setHorizontalDpi(300);
            s.getCTWorksheet().getPageSetup().setVerticalDpi(300);
            assertEquals(600, PrintSetup.read(wb, s, 0, 9, 0).hdpi);
            boolean keep = PrintSetup.useFileDpi;
            try {
                PrintSetup.useFileDpi = true;
                PrintSetup p = PrintSetup.read(wb, s, 0, 9, 0);
                assertEquals(300, p.hdpi);
                assertEquals(300, p.vdpi);
                assertEquals(150, PrintSetup.read(wb, s, 0, 9, 150).hdpi);
            } finally {
                PrintSetup.useFileDpi = keep;
            }
        }
    }
}
