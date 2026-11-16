package xlsx2pdf.render;

import java.awt.Color;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.w3c.dom.Element;

/**
 * タイムライン (tsle:timeslicer) を描く: 枠、見出しと下線、選んだ期間の文字、単位 (年・四半期・月・日)、
 * 年の文字、区画の番号、区画の棒 (選んだ期間は青のグラデーション、ほかは灰色)、下のスクロールの線。
 * 寸法と色は Excel の印刷の実測。
 */
final class TimelinePainter {
    private static final String REL_TIMELINE = "http://schemas.microsoft.com/office/2011/relationships/timeline";
    private static final String REL_TIMELINE_CACHE = "http://schemas.microsoft.com/office/2011/relationships/timelineCache";

    /** 枠の上端からの位置 (pt): 見出しの基線、下線、選んだ期間の文字の基線、年の文字の基線、番号の基線、棒の上端、棒の高さ、スクロールの線。 */
    static final double CAPTION_BASE = 15.1, CAPTION_LINE = 22.9, LABEL_BASE = 39.3, YEAR_BASE = 60.9, NUM_BASE = 80.5,
            BAR_TOP = 87.0, BAR_H = 13.0, SCROLL = 105.0;
    /** 左右の余白、区画の間隔、区画の間の空き (pt)。 */
    static final double INSET = 5.9, PITCH = 24.5, GAP = 1.7;

    private final RenderContext ctx;
    private final XSSFSheet sheet;

    TimelinePainter(RenderContext ctx, XSSFSheet sheet) {
        this.ctx = ctx;
        this.sheet = sheet;
    }

    private static Element parse(PackagePart p) throws Exception {
        try (InputStream in = p.getInputStream()) {
            return Xml.parse(in);
        }
    }

    /** 名前 name のタイムラインを矩形 r に描く。 */
    void paint(PdfCanvas cv, String name, Rectangle2D r) throws Exception {
        Element tl = null;
        PackagePart sp = sheet.getPackagePart();
        for (PackageRelationship rel : sp.getRelationshipsByType(REL_TIMELINE)) {
            for (Element t : Xml.children(parse(sp.getRelatedPart(rel)), "timeline")) if (name.equals(t.getAttribute("name"))) tl = t;
        }
        if (tl == null) return;
        Element cache = null;
        PackagePart wp = sheet.getWorkbook().getPackagePart();
        for (PackageRelationship rel : wp.getRelationshipsByType(REL_TIMELINE_CACHE)) {
            Element c = parse(wp.getRelatedPart(rel));
            if (tl.getAttribute("cache").equals(c.getAttribute("name"))) cache = c;
        }
        Element state = Xml.child(cache, "state");
        Element bounds = Xml.child(state, "bounds"), selection = Xml.child(state, "selection");
        LocalDate b0 = date(bounds, "startDate"), b1 = date(bounds, "endDate");
        if (b0 == null || b1 == null) { b0 = LocalDate.now().withDayOfYear(1); b1 = b0.plusYears(1); }
        LocalDate s0 = date(selection, "startDate"), s1 = date(selection, "endDate");
        int level = Xml.intAttr(tl, "level", 2);
        LocalDate scroll = tl.hasAttribute("scrollPosition") ? LocalDate.parse(tl.getAttribute("scrollPosition").substring(0, 10)) : b0;
        String caption = tl.hasAttribute("caption") ? tl.getAttribute("caption") : name;
        boolean showHeader = Xml.boolAttr(tl, "showHeader", true), showSel = Xml.boolAttr(tl, "showSelectionLabel", true),
                showLevel = Xml.boolAttr(tl, "showTimeLevel", true), showScroll = Xml.boolAttr(tl, "showHorizontalScrollbar", true);
        String style = Xml.attr(tl, "style", "TimeSlicerStyleLight1");
        int n = 1;
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("(\\d+)$").matcher(style);
        if (m.find()) n = Math.max(1, Math.min(6, Integer.parseInt(m.group(1))));
        Color accent = ctx.theme.scheme.getOrDefault("accent" + n, new Color(0x156082));
        Color gray = new Color(0x808080), light = new Color(0xA6A6A6), bar = new Color(0xD9D9D9);

        cv.save();
        cv.clipRect(r.getX(), r.getY(), r.getWidth(), r.getHeight());
        cv.fillRect(r.getX(), r.getY(), r.getWidth(), r.getHeight(), Color.WHITE);
        FontManager.Spec font = ctx.fonts.spec(ctx.theme.minorLatin, ctx.theme.fontRef("+mn-ea", true), false, false);
        FontManager.Spec bold = ctx.fonts.spec(ctx.theme.minorLatin, ctx.theme.fontRef("+mn-ea", true), true, false);
        double x0 = r.getX() + INSET, y0 = r.getY(), w = r.getWidth() - 2 * INSET;
        boolean ja = ctx.japanese;
        if (showHeader) {
            cv.text(caption, x0 + 1.3, y0 + CAPTION_BASE, bold, 11f, Color.BLACK, 1, 0);
            cv.fillRect(x0, y0 + CAPTION_LINE - 0.35, w, 0.7, accent);
        }
        if (showSel) cv.text(selectionLabel(s0, s1, level, ja), x0, y0 + LABEL_BASE, font, 10f, accent, 1, 0);
        if (showLevel) {
            String lv = switch (level) { case 0 -> ja ? "年" : "YEARS"; case 1 -> ja ? "四半期" : "QUARTERS"; case 3 -> ja ? "日" : "DAYS"; default -> ja ? "月" : "MONTHS"; };
            double lw = font.width(lv, 9f, ctx.fonts);
            double tx = r.getMaxX() - INSET - 9 - lw;
            cv.text(lv, tx, y0 + LABEL_BASE, font, 9f, light, 1, 0);
            Path2D tri = new Path2D.Double();
            double cx = r.getMaxX() - INSET - 3.5, cy = y0 + LABEL_BASE - 3;
            tri.moveTo(cx - 2.6, cy - 1.3); tri.lineTo(cx + 2.6, cy - 1.3); tri.lineTo(cx, cy + 1.5); tri.closePath();
            cv.fill(tri, light);
        }
        // 区画: scrollPosition の区画から、枠に収まる数だけ
        int count = (int) Math.max(1, Math.floor((w + GAP) / PITCH));
        LocalDate t = floor(scroll, level);
        int prevYear = -1;
        for (int i = 0; i < count && t.isBefore(b1); i++) {
            LocalDate next = step(t, level);
            double bx = x0 + i * PITCH;
            boolean sel = s0 == null || s1 == null || !t.isAfter(s1) && next.isAfter(s0);
            if (sel) cv.fillAxial(bx, y0 + BAR_TOP, PITCH - GAP, BAR_H, bx, y0 + BAR_TOP, bx, y0 + BAR_TOP + BAR_H,
                    SlicerPainter.tint(accent, 0.6), SlicerPainter.tint(accent, 0.1));
            else cv.fillRect(bx, y0 + BAR_TOP, PITCH - GAP, BAR_H, bar);
            String num = switch (level) {
                case 0 -> String.valueOf(t.getYear());
                case 1 -> (ja ? "第" : "Q") + ((t.getMonthValue() - 1) / 3 + 1) + (ja ? "四半期" : "");
                case 3 -> String.valueOf(t.getDayOfMonth());
                default -> ja ? String.valueOf(t.getMonthValue()) : t.getMonth().getDisplayName(java.time.format.TextStyle.SHORT, java.util.Locale.ENGLISH).toUpperCase(java.util.Locale.ROOT);
            };
            cv.text(num, bx + 2, y0 + NUM_BASE, font, 9f, gray, 1, 0);
            if (level > 0 && t.getYear() != prevYear) {
                String yl = level == 3 ? t.getYear() + (ja ? " 年 " : " ") + t.getMonthValue() + (ja ? " 月" : "") : String.valueOf(t.getYear());
                if (level != 3 || t.getDayOfMonth() == 1 || i == 0) cv.text(yl, bx + 2, y0 + YEAR_BASE, font, 9f, light, 1, 0);
                prevYear = t.getYear();
            }
            t = next;
        }
        if (showScroll) cv.fillRect(x0, y0 + SCROLL, w, 0.7, new Color(0xC4C4C4));
        cv.stroke(new Rectangle2D.Double(r.getX() + 0.35, r.getY() + 0.35, r.getWidth() - 0.7, r.getHeight() - 0.7), new PdfCanvas.Stroke(0.7f), accent);
        cv.restore();
    }

    private static LocalDate date(Element e, String attr) {
        if (e == null || !e.hasAttribute(attr)) return null;
        String v = e.getAttribute(attr);
        return LocalDate.parse(v.length() >= 10 ? v.substring(0, 10) : v);
    }

    /** 単位の区画の始まり。 */
    private static LocalDate floor(LocalDate d, int level) {
        return switch (level) {
            case 0 -> d.withDayOfYear(1);
            case 1 -> LocalDate.of(d.getYear(), ((d.getMonthValue() - 1) / 3) * 3 + 1, 1);
            case 3 -> d;
            default -> d.withDayOfMonth(1);
        };
    }

    private static LocalDate step(LocalDate d, int level) {
        return switch (level) {
            case 0 -> d.plusYears(1);
            case 1 -> d.plusMonths(3);
            case 3 -> d.plusDays(1);
            default -> d.plusMonths(1);
        };
    }

    /** 選んだ期間の文字 (例: 2024 年 3 月 ～ 6 月)。選んでいなければ「すべての期間」。 */
    static String selectionLabel(LocalDate s0, LocalDate s1, int level, boolean ja) {
        if (s0 == null || s1 == null) return ja ? "すべての期間" : "All Periods";
        String a, b;
        switch (level) {
            case 0 -> { a = s0.getYear() + (ja ? " 年" : ""); b = s1.getYear() + (ja ? " 年" : ""); }
            case 1 -> {
                a = s0.getYear() + (ja ? " 年 第 " : " Q") + ((s0.getMonthValue() - 1) / 3 + 1) + (ja ? " 四半期" : "");
                b = (s1.getYear() == s0.getYear() ? "" : s1.getYear() + (ja ? " 年 " : " ")) + (ja ? "第 " : "Q") + ((s1.getMonthValue() - 1) / 3 + 1) + (ja ? " 四半期" : "");
            }
            case 3 -> { a = s0.getYear() + " 年 " + s0.getMonthValue() + " 月 " + s0.getDayOfMonth() + " 日"; b = s1.getMonthValue() + " 月 " + s1.getDayOfMonth() + " 日"; }
            default -> {
                if (!ja) {
                    String ma = s0.getMonth().getDisplayName(java.time.format.TextStyle.SHORT, java.util.Locale.ENGLISH);
                    String mb = s1.getMonth().getDisplayName(java.time.format.TextStyle.SHORT, java.util.Locale.ENGLISH);
                    return s0.getYear() == s1.getYear() ? ma + " - " + mb + " " + s1.getYear() : ma + " " + s0.getYear() + " - " + mb + " " + s1.getYear();
                }
                a = s0.getYear() + " 年 " + s0.getMonthValue() + " 月";
                b = (s1.getYear() == s0.getYear() ? "" : s1.getYear() + " 年 ") + s1.getMonthValue() + " 月";
            }
        }
        if (ChronoUnit.DAYS.between(s0, s1) < 0) return a;
        return a.equals(b) ? a : a + (ja ? " ～ " : " - ") + b;
    }
}
