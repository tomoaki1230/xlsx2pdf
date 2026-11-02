package xlsx2pdf.render;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Element;

/**
 * 「セルに配置」した画像 (Microsoft 365 の機能。セルの値が画像で、vm 属性で値のメタデータを指す)。
 * セルの vm (1 から) → metadata.xml の valueMetadata → futureMetadata の rvb → richData/rdrichvalue.xml の値
 * (構造 _localImage の _rvRel:LocalImageIdentifier) → richData/richValueRel.xml の関係 → 画像、とたどる。
 */
final class CellImages {
    /** vm (1 から) ごとの画像のデータ (画像でない値は null)。 */
    private final List<byte[]> byVm = new ArrayList<>();
    /** vm ごとの画像の名前 (形式の判定用)。 */
    private final List<String> names = new ArrayList<>();

    private static final Map<XSSFWorkbook, CellImages> CACHE = new java.util.WeakHashMap<>();

    /** ブックのセル内画像を読む (無ければ空)。 */
    static synchronized CellImages of(XSSFWorkbook wb) {
        return CACHE.computeIfAbsent(wb, CellImages::new);
    }

    private CellImages(XSSFWorkbook wb) {
        try {
            load(wb.getPackage());
        } catch (Exception e) {
            byVm.clear();
            names.clear();
        }
    }

    private static Element read(OPCPackage pkg, String name) throws Exception {
        PackagePart p = part(pkg, name);
        if (p == null) return null;
        try (InputStream in = p.getInputStream()) {
            return Xml.parse(in);
        }
    }

    private static PackagePart part(OPCPackage pkg, String name) throws Exception {
        var pn = PackagingURIHelper.createPartName(name);
        return pkg.containPart(pn) ? pkg.getPart(pn) : null;
    }

    private void load(OPCPackage pkg) throws Exception {
        Element meta = read(pkg, "/xl/metadata.xml");
        Element rvData = read(pkg, "/xl/richData/rdrichvalue.xml");
        Element rvs = read(pkg, "/xl/richData/rdrichvaluestructure.xml");
        PackagePart relPart = part(pkg, "/xl/richData/richValueRel.xml");
        if (meta == null || rvData == null || rvs == null || relPart == null) return;
        Element rels;
        try (InputStream in = relPart.getInputStream()) {
            rels = Xml.parse(in);
        }
        // 構造ごとの、画像の関係番号のキーの位置
        List<Integer> imgKey = new ArrayList<>();
        for (Element s : Xml.children(rvs, "s")) {
            int k = -1, i = 0;
            for (Element key : Xml.children(s, "k")) {
                if ("_rvRel:LocalImageIdentifier".equals(key.getAttribute("n"))) k = i;
                i++;
            }
            imgKey.add(k);
        }
        // 値ごとの関係番号
        List<Integer> rvRel = new ArrayList<>();
        for (Element rv : Xml.children(rvData, "rv")) {
            int s = Xml.intAttr(rv, "s", -1), rel = -1;
            if (s >= 0 && s < imgKey.size() && imgKey.get(s) >= 0) {
                List<Element> vs = Xml.children(rv, "v");
                int k = imgKey.get(s);
                if (k < vs.size()) {
                    try { rel = Integer.parseInt(Xml.text(vs.get(k)).trim()); } catch (NumberFormatException ignored) { rel = -1; }
                }
            }
            rvRel.add(rel);
        }
        // 関係番号ごとの画像
        List<String> relIds = new ArrayList<>();
        for (Element r : Xml.children(rels, "rel")) relIds.add(r.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"));
        Map<Integer, byte[]> imgData = new HashMap<>();
        Map<Integer, String> imgName = new HashMap<>();
        for (int i = 0; i < relIds.size(); i++) {
            PackageRelationship pr = relPart.getRelationship(relIds.get(i));
            if (pr == null) continue;
            var target = PackagingURIHelper.resolvePartUri(relPart.getPartName().getURI(), pr.getTargetURI());
            PackagePart img = part(pkg, target.getPath());
            if (img == null) continue;
            try (InputStream in = img.getInputStream()) {
                imgData.put(i, in.readAllBytes());
                imgName.put(i, target.getPath());
            }
        }
        // futureMetadata (XLRICHVALUE) の bk ごとの値の番号
        List<Integer> rvb = new ArrayList<>();
        for (Element fm : Xml.children(meta, "futureMetadata")) {
            if (!"XLRICHVALUE".equals(fm.getAttribute("name"))) continue;
            for (Element bk : Xml.children(fm, "bk")) {
                int v = -1;
                Element ext = Xml.path(bk, "extLst", "ext");
                Element b = ext == null ? null : Xml.child(ext, "rvb");
                if (b != null) v = Xml.intAttr(b, "i", -1);
                rvb.add(v);
            }
        }
        Element vmEl = Xml.child(meta, "valueMetadata");
        if (vmEl == null) return;
        for (Element bk : Xml.children(vmEl, "bk")) {
            Element rc = Xml.child(bk, "rc");
            int v = rc == null ? -1 : Xml.intAttr(rc, "v", -1);
            int rv = v >= 0 && v < rvb.size() ? rvb.get(v) : -1;
            int rel = rv >= 0 && rv < rvRel.size() ? rvRel.get(rv) : -1;
            byVm.add(rel >= 0 ? imgData.get(rel) : null);
            names.add(rel >= 0 ? imgName.get(rel) : null);
        }
    }

    /** vm (1 から) の画像のデータ (無ければ null)。 */
    byte[] image(long vm) {
        int i = (int) vm - 1;
        return i >= 0 && i < byVm.size() ? byVm.get(i) : null;
    }

    /** vm (1 から) の画像の名前。 */
    String name(long vm) {
        int i = (int) vm - 1;
        return i >= 0 && i < names.size() ? names.get(i) : null;
    }
}
