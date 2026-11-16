package xlsx2pdf.render;

import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.util.Map;

import org.w3c.dom.Element;

/**
 * 図形の 3D の回転 (spPr/scene3d/camera) と奥行き (sp3d/extrusionH)。回転は図形の中心を軸に、左右の回転 X、上下の傾き Y、
 * 面の中の回転 Z の順にかける。Excel の印刷の遠近はごく弱いので、平行投影 (アフィン変換) で近似する。
 */
final class Scene3D {
    private Scene3D() {}

    /** カメラの既定値ごとの回転 (Excel の RotationX・Y・Z。度)。Excel 2024 の COM で読み出した値。 */
    private static final Map<String, double[]> CAMERAS = Map.ofEntries(
            Map.entry("isometricTopUp", new double[] {45.2809, -35.39627, -60.16242}),
            Map.entry("isometricTopDown", new double[] {-45.28088, -35.39627, 60.16245}),
            Map.entry("isometricBottomUp", new double[] {-45.28088, 35.39627, -60.16242}),
            Map.entry("isometricBottomDown", new double[] {45.2809, 35.39627, 60.16245}),
            Map.entry("isometricLeftUp", new double[] {-45, -35, 0}),
            Map.entry("isometricLeftDown", new double[] {-45, 35, 0}),
            Map.entry("isometricRightUp", new double[] {45, 35, 0}),
            Map.entry("isometricRightDown", new double[] {45, -35, 0}),
            Map.entry("isometricOffAxis1Left", new double[] {-64, 18, 0}),
            Map.entry("isometricOffAxis1Right", new double[] {26, 18, 0}),
            Map.entry("isometricOffAxis1Top", new double[] {53.45425, -58.73812, -57.6425}),
            Map.entry("isometricOffAxis2Left", new double[] {-26, 18, 0}),
            Map.entry("isometricOffAxis2Right", new double[] {64, 18, 0}),
            Map.entry("isometricOffAxis2Top", new double[] {-53.45424, -58.73812, 57.6425}),
            Map.entry("isometricOffAxis3Left", new double[] {-64, -18, 0}),
            Map.entry("isometricOffAxis3Right", new double[] {26, -18, 0}),
            Map.entry("isometricOffAxis3Bottom", new double[] {53.45425, 58.73808, 57.64255}),
            Map.entry("isometricOffAxis4Left", new double[] {-26, -18, 0}),
            Map.entry("isometricOffAxis4Right", new double[] {64, -18, 0}),
            Map.entry("isometricOffAxis4Bottom", new double[] {-53.45424, 58.73808, -57.64252}),
            Map.entry("perspectiveLeft", new double[] {-20, 0, 0}),
            Map.entry("perspectiveRight", new double[] {20, 0, 0}),
            Map.entry("perspectiveAbove", new double[] {0, -20, 0}),
            Map.entry("perspectiveBelow", new double[] {0, 20, 0}),
            Map.entry("perspectiveAboveLeftFacing", new double[] {-14.33758, 39.30647, 18.87617}),
            Map.entry("perspectiveAboveRightFacing", new double[] {14.33762, 39.30647, -18.87615}),
            Map.entry("perspectiveContrastingLeftFacing", new double[] {-43.93887, 10.39642, 3.553517}),
            Map.entry("perspectiveContrastingRightFacing", new double[] {43.9389, 10.39642, -3.553517}),
            Map.entry("perspectiveHeroicLeftFacing", new double[] {-14.29302, -11.01627, -2.626817}),
            Map.entry("perspectiveHeroicRightFacing", new double[] {14.29307, -11.01627, 2.626767}),
            Map.entry("perspectiveHeroicExtremeLeftFacing", new double[] {-34.46068, 8.12245, 2.908583}),
            Map.entry("perspectiveHeroicExtremeRightFacing", new double[] {34.46073, 8.12245, -2.9086}),
            Map.entry("perspectiveRelaxed", new double[] {0, -50.43993, 0}),
            Map.entry("perspectiveRelaxedModerately", new double[] {0, -35.15602, 0}));

    /** 斜投影のカメラの、奥行きの向き (奥の面のずれ。奥行き 1 あたり)。 */
    static final double OBLIQUE = 0.25;

    /** 回転と奥行き。 */
    record View(AffineTransform at, double depthX, double depthY) {
        boolean extruded() { return depthX != 0 || depthY != 0; }
    }

    /**
     * 図形の 3D の見え方。回転も奥行きもなければ null。
     *
     * @param r 図形の矩形 (回転の中心)
     */
    static View of(Element spPr, Rectangle2D r) {
        Element scene = Xml.child(spPr, "scene3d");
        Element cam = Xml.child(scene, "camera");
        Element sp3d = Xml.child(spPr, "sp3d");
        double depth = sp3d == null ? 0 : Xml.longAttr(sp3d, "extrusionH", 0) / 12700.0;
        if (cam == null) return null;
        String prst = cam.getAttribute("prst");
        double[] xyz = CAMERAS.getOrDefault(prst, new double[3]);
        Element rot = Xml.child(cam, "rot");
        if (rot != null) {
            xyz = new double[] {-signed(Xml.longAttr(rot, "lon", 0) / 60000.0), signed(Xml.longAttr(rot, "lat", 0) / 60000.0),
                    -signed(Xml.longAttr(rot, "rev", 0) / 60000.0)};
        }
        double a = Math.toRadians(xyz[0]), b = Math.toRadians(xyz[1]), c = Math.toRadians(xyz[2]);
        double ca = Math.cos(a), sa = Math.sin(a), cb = Math.cos(b), sb = Math.sin(b), cc = Math.cos(c), sc = Math.sin(c);
        // 単位ベクトルの投影 (x 右、y 下、z 手前)
        double xx = ca * cc + sa * sb * sc, xy = ca * sc - sa * sb * cc;
        double yx = -cb * sc, yy = cb * cc;
        double zx = sa * cc - ca * sb * sc, zy = sa * sc + ca * sb * cc;
        double dx = -depth * zx, dy = -depth * zy;
        if (prst.startsWith("oblique") || prst.startsWith("legacyOblique") || prst.startsWith("legacyPerspective")) {
            // 斜投影: 面は回さず、奥の面を名前の向きへずらす
            String d = prst.replace("legacy", "").replace("Oblique", "").replace("Perspective", "").replace("oblique", "");
            double ox = d.contains("Left") ? -1 : d.contains("Right") ? 1 : 0, oy = d.contains("Top") ? -1 : d.contains("Bottom") ? 1 : 0;
            dx = depth * OBLIQUE * ox;
            dy = depth * OBLIQUE * oy;
        }
        boolean rotated = Math.abs(xx - 1) > 1e-6 || Math.abs(xy) > 1e-6 || Math.abs(yx) > 1e-6 || Math.abs(yy - 1) > 1e-6;
        if (!rotated && dx == 0 && dy == 0) return null;
        AffineTransform at = new AffineTransform();
        at.translate(r.getCenterX(), r.getCenterY());
        at.concatenate(new AffineTransform(xx, xy, yx, yy, 0, 0));
        at.translate(-r.getCenterX(), -r.getCenterY());
        return new View(at, dx, dy);
    }

    private static double signed(double deg) {
        double d = deg % 360;
        return d > 180 ? d - 360 : d;
    }
}
