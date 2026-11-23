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
 * zip の中の XML をメモリ上で書き換える。
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

    /** 判定で読む先頭の大きさ (Strict の名前空間は先頭の要素にある)。 */
    static final int PREFIX = 1 << 20;

    /** 先頭の max バイトまでを読む (それより大きい項目でも例外にしない)。 */
    private static byte[] readPrefix(InputStream in, int max) throws IOException {
        return in.readNBytes(max);
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

    /** ファイルの zip に Strict の名前空間があるか (目次から 2 つの項目だけを読む。ほかの項目は展開しない)。 */
    static boolean isStrict(java.io.File f, long maxEntryBytes) throws IOException {
        try (java.util.zip.ZipFile z = new java.util.zip.ZipFile(f)) {
            for (String n : new String[] {"_rels/.rels", "xl/workbook.xml"}) {
                ZipEntry e = z.getEntry(n);
                if (e == null) continue;
                try (InputStream in = z.getInputStream(e)) {
                    byte[] b = readPrefix(in, PREFIX);
                    if (indexOf(b, STRICT_BYTES) >= 0) return true;
                }
            }
        }
        return false;
    }

    /**
     * メモリの zip に Strict の名前空間があるか。目次 (中央ディレクトリ) から 2 つの項目の位置を求め、その項目だけを展開する
     * (ほかの項目は展開しない)。目次を読めなければ Strict でないとみなす (POI が開くときに調べる)。
     */
    static boolean isStrict(byte[] zip, long maxEntryBytes) throws IOException {
        java.nio.ByteBuffer b = java.nio.ByteBuffer.wrap(zip).order(java.nio.ByteOrder.LITTLE_ENDIAN);
        int eocd = -1;
        for (int i = zip.length - 22; i >= Math.max(0, zip.length - 22 - 65535); i--) {
            if (b.getInt(i) == 0x06054b50) { eocd = i; break; }
        }
        if (eocd < 0) return false;
        int count = b.getShort(eocd + 10) & 0xFFFF;
        long cd = b.getInt(eocd + 16) & 0xFFFFFFFFL;
        int p = (int) Math.min(cd, zip.length);
        for (int k = 0; k < count && p + 46 <= zip.length; k++) {
            if (b.getInt(p) != 0x02014b50) break;
            int nameLen = b.getShort(p + 28) & 0xFFFF, extraLen = b.getShort(p + 30) & 0xFFFF, commentLen = b.getShort(p + 32) & 0xFFFF;
            long local = b.getInt(p + 42) & 0xFFFFFFFFL;
            if (p + 46 + nameLen > zip.length) break;
            String n = new String(zip, p + 46, nameLen, java.nio.charset.StandardCharsets.UTF_8);
            if ((n.equals("_rels/.rels") || n.equals("xl/workbook.xml")) && local < zip.length) {
                // その項目の局所ヘッダーから読む (ZipInputStream は最初の項目だけを展開する)
                try (ZipInputStream zi = new ZipInputStream(new ByteArrayInputStream(zip, (int) local, zip.length - (int) local))) {
                    ZipEntry e = zi.getNextEntry();
                    if (e != null && e.getName().equals(n)) {
                        byte[] d = readPrefix(zi, PREFIX);
                        if (indexOf(d, STRICT_BYTES) >= 0) return true;
                    }
                }
            }
            p += 46 + nameLen + extraLen + commentLen;
        }
        return false;
    }

    /** バイト列が Strict なら書き換え、そうでなければ null。 */
    static byte[] convertIfStrict(byte[] book, long maxEntryBytes) throws IOException {
        if (!isStrict(book, maxEntryBytes)) return null;
        return toTransitional(new ByteArrayInputStream(book), maxEntryBytes, maxEntryBytes > 0 ? 1L << 30 : 0);
    }
}
