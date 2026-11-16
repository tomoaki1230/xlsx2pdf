package xlsx2pdf.render;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.awt.Color;
import java.awt.geom.AffineTransform;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;

import org.junit.jupiter.api.Test;

class SvgPainterTest {
    @Test
    void parsesCompactNumbers() {
        assertArrayEquals(new double[] {1, -2.5, 0.5, 3e2}, SvgPainter.numbers("1-2.5.5,3e2"), 1e-9);
    }

    @Test
    void parsesColors() {
        assertEquals(new Color(0x112233), SvgPainter.color("#123", null));
        assertEquals(new Color(0xC00000), SvgPainter.color("#C00000", null));
        assertEquals(new Color(112, 173, 71), SvgPainter.color("rgb(112, 173, 71)", null));
        assertEquals(new Color(0xFFA500), SvgPainter.color("orange", null));
        assertEquals(new Color(0x808080), SvgPainter.color("grey", null));
        assertEquals(Color.RED, SvgPainter.color("currentColor", Color.RED));
        assertNull(SvgPainter.color("none", Color.RED));
    }

    @Test
    void parsesTransforms() {
        AffineTransform t = SvgPainter.transform("translate(10 20) scale(2)");
        Point2D p = t.transform(new Point2D.Double(1, 1), null);
        assertEquals(12, p.getX(), 1e-9);
        assertEquals(22, p.getY(), 1e-9);
        Point2D q = SvgPainter.transform("rotate(90 5 5)").transform(new Point2D.Double(10, 5), null);
        assertEquals(5, q.getX(), 1e-9);
        assertEquals(10, q.getY(), 1e-9);
    }

    @Test
    void parsesPathsWithRelativeCommandsAndArcs() {
        Rectangle2D b = SvgPainter.SvgPath.parse("m60 10 h40 v40 h-40 z").getBounds2D();
        assertEquals(60, b.getX(), 1e-9);
        assertEquals(40, b.getWidth(), 1e-9);
        assertEquals(40, b.getHeight(), 1e-9);
        // 半円 (0,0) → (20,0)、半径 10、sweep=1 は画面で時計回り = 上 (y の負の側) を通る
        Path2D arc = SvgPainter.SvgPath.parse("M0 0 A10 10 0 0 1 20 0");
        Rectangle2D a = arc.getBounds2D();
        assertEquals(20, a.getWidth(), 1e-6);
        assertEquals(-10, a.getY(), 0.5);
        assertEquals(20, arc.getCurrentPoint().getX(), 1e-9);
        assertEquals(0, arc.getCurrentPoint().getY(), 1e-9);
    }
}
