package xlsx2pdf;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * Strict OOXML (ISO/IEC 29500 Strict。Excel の「Strict Open XML スプレッドシート」) のブックを、POI が読める通常 (Transitional) の形に直す。
 * 違いは名前空間とリレーションの種類の URI (http://purl.oclc.org/ooxml/... → http://schemas.openxmlformats.org/.../2006/...) なので、
 * zip の中の XML をメモリ上で書き換える。夜間 n3415・n5623。
 */
final class StrictOoxml {
    private StrictOoxml() {}

    private static final String STRICT = "http://purl.oclc.org/ooxml/";
    private static final byte[] STRICT_BYTES = STRICT.getBytes(StandardCharsets.US_ASCII);

    /** 名前の変わる URI (先に置き換える)。 */
    private static final String[][] SPECIAL = {
        {"http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"},
        {"http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"},
        {"http://purl.oclc.org/ooxml/officeDocument/extendedProperties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"},
        {"http://purl.oclc.org/ooxml/officeDocument/customProperties", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"},
    };
    private static final java.util.regex.Pattern GENERIC = java.util.regex.Pattern.compile("http://purl\\.oclc\\.org/ooxml/(\\w+)/");

    /** zip の先頭の項目に Strict の名前空間があるか (_rels/.rels か [Content_Types].xml の次に来る XML を見る)。 */
    static boolean isStrict(InputStream zip, long maxEntryBytes) throws IOException {
        try (ZipInputStream zi = new ZipInputStream(zip)) {
            ZipEntry e;
            int seen = 0;
            while ((e = zi.getNextEntry()) != null && seen < 64) {
                String n = e.getName();
                if (!(n.equals("_rels/.rels") || n.equals("xl/workbook.xml"))) continue;
                seen++;
                byte[] b = readLimited(zi, Math.min(maxEntryBytes > 0 ? maxEntryBytes : Long.MAX_VALUE, 4L << 20));
                if (indexOf(b, STRICT_BYTES) >= 0) return true;
                if (n.equals("xl/workbook.xml")) return false;
            }
        }
        return false;
    }

    /**
     * Strict のブックを通常の形に書き換えた zip を返す。
     *
     * @param maxEntryBytes 1 項目の展開後の上限 (0 以下は無制限)
     * @param maxTotalBytes 全体の展開後の上限 (0 以下は無制限)
     */
    static byte[] toTransitional(InputStream zip, long maxEntryBytes, long maxTotalBytes) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        long total = 0;
        try (ZipInputStream zi = new ZipInputStream(zip); ZipOutputStream zo = new ZipOutputStream(out)) {
            ZipEntry e;
            while ((e = zi.getNextEntry()) != null) {
                if (e.isDirectory()) continue;
                byte[] b = readLimited(zi, maxEntryBytes > 0 ? maxEntryBytes : Long.MAX_VALUE);
                total += b.length;
                if (maxTotalBytes > 0 && total > maxTotalBytes) throw new IOException("Strict OOXML のブックが大きすぎます (展開後 " + total + " バイト超)");
                String n = e.getName().toLowerCase(java.util.Locale.ROOT);
                if ((n.endsWith(".xml") || n.endsWith(".rels") || n.endsWith(".vml")) && indexOf(b, STRICT_BYTES) >= 0) {
                    String s = new String(b, StandardCharsets.UTF_8);
                    for (String[] sp : SPECIAL) s = s.replace(sp[0], sp[1]);
                    s = GENERIC.matcher(s).replaceAll("http://schemas.openxmlformats.org/$1/2006/");
                    b = s.getBytes(StandardCharsets.UTF_8);
                }
                zo.putNextEntry(new ZipEntry(e.getName()));
                zo.write(b);
                zo.closeEntry();
            }
        }
        return out.toByteArray();
    }

    private static byte[] readLimited(InputStream in, long max) throws IOException {
        ByteArrayOutputStream o = new ByteArrayOutputStream();
        byte[] buf = new byte[65536];
        long n = 0;
        int r;
        while ((r = in.read(buf)) > 0) {
            n += r;
            if (n > max) throw new IOException("zip の項目が大きすぎます (" + max + " バイト超)");
            o.write(buf, 0, r);
        }
        return o.toByteArray();
    }

    private static int indexOf(byte[] a, byte[] p) {
        outer:
        for (int i = 0; i + p.length <= a.length; i++) {
            for (int j = 0; j < p.length; j++) if (a[i + j] != p[j]) continue outer;
            return i;
        }
        return -1;
    }

    /** バイト列が Strict なら書き換え、そうでなければ null。 */
    static byte[] convertIfStrict(byte[] book, long maxEntryBytes) throws IOException {
        if (!isStrict(new ByteArrayInputStream(book), maxEntryBytes)) return null;
        return toTransitional(new ByteArrayInputStream(book), maxEntryBytes, 1L << 30);
    }
}
