package xlsx2pdf;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.AccessDeniedException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ThreadLocalRandom;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.cos.COSBase;
import org.apache.pdfbox.cos.COSDictionary;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.cos.COSStream;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChartSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import xlsx2pdf.render.CellPainter;
import xlsx2pdf.render.CondFormat;
import xlsx2pdf.render.DrawingPainter;
import xlsx2pdf.render.HeaderFooter;
import xlsx2pdf.render.Paginator;
import xlsx2pdf.render.PdfCanvas;
import xlsx2pdf.render.PrintSetup;
import xlsx2pdf.render.RenderContext;
import xlsx2pdf.render.FontManager;
import xlsx2pdf.render.SheetLayout;

/** xlsx ブックを、Excel が印刷するのと同じように PDF に変換する。 */
public final class Converter {
    /** 図形・グラフの切り抜きをブロックの外側に広げる幅 (用紙上の pt。セルは罫線のため 2pt)。 */
    static final double DRAWING_CLIP = 0.44;

    /** 変換のオプション。 */
    public static final class Options {
        /** 追加で検索するフォントフォルダ。 */
        public final List<File> fontDirs = new ArrayList<>();
        /**
         * フォントの走査結果のキャッシュ (fonts-*.tsv) を置くフォルダ。null (既定) なら ~/.cache/xlsx2pdf。
         * フォルダがなければ作る。
         */
        public File fontCacheDir;
        /** フォントの走査結果のキャッシュを使わない (読まず、書かない)。変換のたびにフォントファイルを読むので遅くなる。 */
        public boolean noFontCache;
        /**
         * 数式を再計算してから印刷する。既定 (false) は、ファイルに保存された計算結果 (最後に保存したときの値) を使う。
         * 計算できない数式は保存された結果のまま ({@link xlsx2pdf.render.Recalc})。
         */
        public boolean recalc;
        /** 実際の用紙サイズではなく、Excel の PDF のページサイズ (用紙サイズ ÷ 印刷倍率) を使う。 */
        public boolean excelPageSize;
        /** 印刷するシート名 (空ならすべての表示シート)。 */
        public final List<String> sheets = new ArrayList<>();
        /** 暗号化されたブックのパスワード (Excel の既定の "VelvetSweatshop" は常に試す)。 */
        public String password;
        /** 選択されたシート (tabSelected) だけを印刷する (Excel の既定「作業中のシート」と同じ)。 */
        public boolean activeOnly;
        /** シートに用紙サイズがないときの SpreadsheetML の用紙コード (0 = 日本語のブックは A4、それ以外は Letter)。 */
        public int defaultPaper;
        /**
         * 出力環境: PDF を出力する PC の画面 DPI (96 = 100% 表示、120 = 125%、144 = 150%、192 = 200%)。
         * Excel はこの画面の値で列幅・行高を換算する。0 = 列幅から推定する。
         */
        public int screenDpi;
        /** 出力環境: Excel で選択しているプリンタの解像度 (Microsoft Print to PDF は 600)。0 = 600 dpi。 */
        public int printerDpi;
        /**
         * 作成者の環境に合わせる (--author-env): 作成者の Excel の版・画面 DPI をファイルから推測し、作成者が出力した PDF に近づける。
         * false (既定) は、今の Excel を表示倍率 100% (または {@link #screenDpi}) の PC で印刷した結果に合わせる。
         */
        public boolean authorEnv;

        // ---- 上限 (悪意のある小さなファイルでメモリや時間を使い切られないため。0 以下は無制限) ----
        /**
         * ページ分割で作るページ数 (白紙のページを除く前、全シートの合計) の上限。使用範囲がシートの端まで広がった
         * ファイル (3KB で数千万ページになる) を防ぐ。既定 10,000。
         */
        public long maxPages = 10_000;
        /** 1 枚の画像 (埋め込み画像・EMF の中のビットマップ) の画素数の上限。既定 5,000 万 (約 7,000 × 7,000)。 */
        public long maxImagePixels = 50_000_000L;
        /** xlsx (zip) の 1 項目の展開後の大きさの上限 (バイト)。既定 1 GB。 */
        public long maxZipEntryBytes = 1L << 30;
        /** xlsx (zip) の項目の数の上限。既定 100,000 (POI の既定 1,000 は、数百のシートがある実際のブックで足りない)。0 以下は無制限。 */
        public long maxZipEntries = 100_000;
        /** xlsx (zip) の項目の圧縮率の下限 (展開後に対する圧縮後の比。小さすぎるものは zip 爆弾とみなす)。既定 0.005 (POI の既定 0.01 では、Excel が開く書式の多いブックの styles.xml (先頭の比 0.008)。 */
        public double minInflateRatio = 0.005;
        /** 変換の時間の上限 (秒)。既定 0 = 無制限。 */
        public long timeoutSeconds;
        /** ストリームの入力の大きさの上限 (バイト)。0 以下は無制限 (配列の大きさの限界まで)。 */
        public long maxInputBytes = 1L << 30;
        /** ストリームから変換するときのブック名 (ヘッダーの &amp;F と PDF のタイトル)。null なら空。 */
        public String sourceName;
        /**
         * PDF の編集を禁止する (内容の変更、注釈・フォームの記入、ページの組み替え)。印刷と文字のコピーは許可する。
         * PDF を開くパスワードは付けない。制限はビューアーが守ることが前提で、改ざんを確実に防ぐものではない。
         */
        public boolean noEdit;
        /** 編集の禁止を解除する権限パスワード。null か空なら、知らせない乱数のパスワードにする。 */
        public String ownerPassword;

        /** すべての上限を外す (メモリが潤沢で、時間を掛けてでも変換したいとき)。 */
        public void noLimits() {
            maxPages = 0;
            maxImagePixels = 0;
            maxZipEntryBytes = 0;
            maxZipEntries = 0;
            minInflateRatio = 0;
            timeoutSeconds = 0;
            maxInputBytes = 0;
        }
    }

    /** これまでに作ったページ数 (白紙を除く前、ページ数の上限の判定用)。 */
    private long pagesMade;

    /** 変換の時間の期限 (System.nanoTime。0 = 無制限)。 */
    private long deadline;

    /** 最後の変換の警告。 */
    private List<String> warnings = List.of();
    /** 変換中の警告の集め先。 */
    private xlsx2pdf.render.Warnings collecting;

    /**
     * 最後の変換の警告の一覧 (描けなかったセル・図形・画像など)。変換は警告があっても最後まで続ける。
     *
     * @return 警告の文の一覧 (なければ空)
     */
    public synchronized List<String> getWarnings() {
        return warnings;
    }

    /** 時間の上限を超えていれば変換を中止する。 */
    void checkDeadline() {
        if (deadline != 0 && System.nanoTime() - deadline > 0) {
            throw new LimitExceededException("変換の時間の上限 (" + opt.timeoutSeconds + " 秒) を超えました。--timeout で延ばすか 0 で無制限にできます", true);
        }
    }

    /** はみ出しを調べる列数の上限。 */
    static final int OV_RANGE_COLS = 200;

    /**
     * ToUnicode の 1 行: "&lt;src&gt; &lt;dst&gt;" (bfchar) または "&lt;src&gt; &lt;src&gt; &lt;dst&gt;" (PDFBox が書く 1 コードの bfrange)
     * で、dst が U+2E80〜U+2FDF (部首・康煕部首) または U+FE10〜U+FE19・U+FE30〜U+FE44 (縦書き用の形) のもの。
     */
    private static final Pattern RADICAL_ENTRY = Pattern.compile(
            "^(\\s*(?:<([0-9A-Fa-f]+)>\\s*<\\2>|<[0-9A-Fa-f]+>)\\s*)<(2E[89A-Fa-f][0-9A-Fa-f]|2F[0-9A-Da-d][0-9A-Fa-f]|FE1[0-9]|FE3[0-9A-Fa-f]|FE4[0-4])>(\\s*)$");
    /**
     * 縦書き用の形 (U+FE10〜) に対する横書きの文字。フォントが「：」と「︓」に同じグリフを使うと、PDFBox は小さい方の
     * 「︓」を ToUnicode に書くため、コピーした文字が縦書き用の形になる。
     */
    private static final String VERTICAL_FORM_SRC = "\uFE10\uFE11\uFE12\uFE13\uFE14\uFE15\uFE16\uFE17\uFE18\uFE19"
            + "\uFE30\uFE31\uFE32\uFE33\uFE34\uFE35\uFE36\uFE37\uFE38\uFE39\uFE3A\uFE3B\uFE3C\uFE3D\uFE3E\uFE3F"
            + "\uFE40\uFE41\uFE42\uFE43\uFE44";
    private static final String VERTICAL_FORM_DST = "\uFF0C\u3001\u3002\uFF1A\uFF1B\uFF01\uFF1F\u3016\u3017\u2026"
            + "\u2025\u2014\u2013\uFF3F\uFF3F\uFF08\uFF09\uFF5B\uFF5D\u3014\u3015\u3010\u3011\u300A\u300B\u3008"
            + "\u3009\u300C\u300D\u300E\u300F";

    private final Options opt;

    /** 指定のオプションで変換器を作る。 */
    public Converter(Options opt) {
        this.opt = opt;
    }

    /**
     * 1 シート分の印刷の準備結果。
     *
     * @param index    ブック内のシート番号
     * @param layout   レイアウト (グラフシートは null)
     * @param pages    印刷するページ
     * @param drawings 図形・画像・グラフ
     * @param chart    グラフシートか
     */
    private record SheetJob(XSSFSheet sheet, int index, PrintSetup ps, SheetLayout layout, List<Paginator.Page> pages,
                            DrawingPainter drawings, boolean chart) {}

    /**
     * ブックを PDF に変換する。
     *
     * @param in  入力の xlsx
     * @param out 出力の PDF
     * @throws IOException            読み書きに失敗したとき、または対応していない形式・パスワードのとき
     * @throws InvalidFormatException xlsx として不正なとき
     */
    public synchronized void convert(File in, File out) throws IOException, InvalidFormatException {
        if (samePath(in, out)) throw new IOException("入力と出力が同じファイルです: " + out);
        Path target = writeTarget(out);
        // デバイスなど通常のファイルでない出力先は置き換えられないので、直接書く
        if (Files.exists(target) && !Files.isRegularFile(target)) {
            render(in, out);
            return;
        }
        Path tmp;
        try {
            tmp = uniqueTemp(target);
        } catch (NoSuchFileException e) {
            throw new IOException("出力先のフォルダがありません: " + target.getParent());
        } catch (AccessDeniedException e) {
            // フォルダに書けなくても、既存のファイルに書ける場合は直接書く
            if (!Files.isWritable(target)) throw new IOException("出力先に書き込めません: " + target);
            render(in, out);
            return;
        }
        try {
            copyPermissions(target, tmp);
            render(in, tmp.toFile());
            try {
                moveOver(tmp, target);
                if (xlsx2pdf.render.ValueDump.ON) {   // 値の書き出しも出力の名前に合わせる
                    Path v = Path.of(tmp + ".values.tsv");
                    if (Files.exists(v)) Files.move(v, Path.of(target + ".values.tsv"), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                }
            } catch (FileSystemException e) {
                throw new IOException("出力先に書き込めません (ほかのアプリで開いていないか確認してください): " + target);
            }
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    /**
     * ストリームからストリームへ変換する。一時ファイルは作らず、メモリの中で PDF を完成させてから out に書く
     * (失敗したときは out に何も書かない)。in と out は閉じない。ブック名は {@link Options#sourceName} で渡す。
     * 同じ Converter を複数のスレッドで同時に呼ぶと順番に処理する。並列に変換するときはスレッドごとに Converter を作る。
     *
     * @param in  入力の xlsx
     * @param out 出力の PDF
     * @throws IOException            読み書きに失敗したとき、または対応していない形式・パスワードのとき
     * @throws InvalidFormatException xlsx として不正なとき
     * @throws LimitExceededException 入力が {@link Options#maxInputBytes} を超えたとき、ほかの上限を超えたとき
     */
    public synchronized void convert(InputStream in, OutputStream out) throws IOException, InvalidFormatException {
        java.util.Objects.requireNonNull(in, "in");
        java.util.Objects.requireNonNull(out, "out");
        byte[] book = readLimited(in);
        java.io.ByteArrayOutputStream pdf = new java.io.ByteArrayOutputStream();
        prepare();
        String name = opt.sourceName != null ? opt.sourceName : "";
        try (OPCPackage pkg = openPackage(book, name)) {
            render(pkg, null, opt.sourceName, doc -> doc.save(pdf));
        }
        byte[] done = finish(pdf.toByteArray());
        out.write(done);
        out.flush();
    }

    /** 入力を上限 {@link Options#maxInputBytes} までメモリに読む。 */
    private byte[] readLimited(InputStream in) throws IOException {
        long limit = opt.maxInputBytes > 0 ? Math.min(opt.maxInputBytes, Integer.MAX_VALUE - 9L) : Integer.MAX_VALUE - 9L;
        byte[] b = in.readNBytes((int) limit + 1);
        if (b.length > limit) throw new LimitExceededException("入力が大きすぎます (上限 " + limit + " バイト)", false);
        return b;
    }

    /** 入力と出力が同じファイルか (別名のパス・リンクも含む)。 */
    static boolean samePath(File in, File out) {
        try {
            return out.exists() && Files.isSameFile(in.toPath(), out.toPath());
        } catch (IOException e) {
            return false;
        }
    }

    /** 実際に書き込むパス。シンボリックリンクはリンク自体を残すため、指す先にする。 */
    private static Path writeTarget(File out) throws IOException {
        Path p = out.toPath().toAbsolutePath();
        return Files.isSymbolicLink(p) && Files.exists(p) ? p.toRealPath() : p;
    }

    /** 同じフォルダに重複しない名前の一時ファイルを作る。権限は通常のファイルの作成と同じ (createTempFile は本人だけになる)。 */
    private static Path uniqueTemp(Path target) throws IOException {
        while (true) {
            String name = target.getFileName() + "." + Long.toHexString(ThreadLocalRandom.current().nextLong()) + ".tmp";
            try {
                return Files.createFile(target.resolveSibling(name));
            } catch (FileAlreadyExistsException e) {
                // 別の名前で作り直す
            }
        }
    }

    /** 既存の出力の権限を一時ファイルに写す (置き換えても権限が変わらないように)。POSIX でない環境では何もしない。 */
    private static void copyPermissions(Path from, Path to) throws IOException {
        if (!Files.exists(from)) return;
        try {
            Files.setPosixFilePermissions(to, Files.getPosixFilePermissions(from));
        } catch (UnsupportedOperationException e) {
            // POSIX の権限のないファイルシステム
        }
    }

    /** 完成した一時ファイルで出力を置き換える。 */
    private static void moveOver(Path tmp, Path target) throws IOException {
        try {
            Files.move(tmp, target, StandardCopyOption.ATOMIC_MOVE);
        } catch (AtomicMoveNotSupportedException e) {
            Files.move(tmp, target, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    /** ブックを読み、PDF を out に書く。 */
    private void render(File in, File out) throws IOException, InvalidFormatException {
        prepare();
        xlsx2pdf.render.ValueDump.begin();
        try (OPCPackage pkg = openPackage(in)) {
            render(pkg, in, in.getName(), doc -> {
                // 読み直せない出力先 (デバイスなど) は、ここで編集を禁止する
                if (out.exists() && !Files.isRegularFile(out.toPath())) protect(doc);
                doc.save(out);
            });
        }
        if (Files.isRegularFile(out.toPath())) finish(out);   // デバイスなどは読み直せない
        if (xlsx2pdf.render.ValueDump.ON)
            Files.write(new File(out.getPath() + ".values.tsv").toPath(), xlsx2pdf.render.ValueDump.take(), java.nio.charset.StandardCharsets.UTF_8);
    }

    /** 完成した PDF の保存先。 */
    private interface PdfSink {
        void save(PDDocument doc) throws IOException;
    }

    /** 変換の前の準備: 上限の設定と、数え直し。 */
    private void prepare() {
        applyLimits();
        pagesMade = 0;
        warnings = List.of();
        deadline = opt.timeoutSeconds > 0 ? System.nanoTime() + opt.timeoutSeconds * 1_000_000_000L : 0;
        if (deadline == 0 && opt.timeoutSeconds > 0) deadline = 1;  // nanoTime がちょうど 0 のとき
    }

    /**
     * 開いたブックを PDF にして sink に渡す。
     *
     * @param source ファイルから読むときの入力 (ストリームなら null)
     * @param name   ブック名 (&amp;F と PDF のタイトル。なければ null)
     */
    private void render(OPCPackage pkg, File source, String name, PdfSink sink) throws IOException {
        collecting = xlsx2pdf.render.Warnings.start();
        try {
            try (XSSFWorkbook wb = new XSSFWorkbook(pkg);
                 PDDocument doc = new PDDocument()) {
                RenderContext ctx = new RenderContext(wb, doc, opt.fontDirs,
                        opt.noFontCache ? null : opt.fontCacheDir != null ? opt.fontCacheDir : FontManager.defaultCacheDir(), source);
                ctx.sourceName = name;
                ctx.deadline = this::checkDeadline;
                ctx.printer.authorEnv = opt.authorEnv;
                ctx.fonts.checkFallbacks();
                if (opt.recalc) {
                    xlsx2pdf.render.Recalc.run(wb, this::checkDeadline);
                    ctx.recalc = true;
                }
                List<SheetJob> jobs = new ArrayList<>();
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    XSSFSheet sheet = wb.getSheetAt(i);
                    // 非表示のシートは出力しない。ただし --sheet で名前を指定したときは出力する
                    boolean named = opt.sheets.contains(sheet.getSheetName());
                    if ((wb.isSheetHidden(i) || wb.isSheetVeryHidden(i)) && !named) continue;
                    if (!opt.sheets.isEmpty() && !named) continue;
                    if (opt.activeOnly && !sheet.isSelected()) continue;
                    xlsx2pdf.render.Warnings.sheet(sheet.getSheetName());
                    SheetJob job = prepare(ctx, sheet, i);
                    if (job != null) jobs.add(job);
                }
                int total = 0;
                // コメントを末尾に印刷するシート (cellComments="atEnd") は、コメントの一覧のページを足す
                java.util.Map<SheetJob, List<List<String[]>>> notes = new java.util.HashMap<>();
                for (SheetJob j : jobs) {
                    total += j.pages.size();
                    if (!j.chart && "atEnd".equals(j.ps.comments)) {
                        List<List<String[]>> np = commentPages(ctx, j);
                        if (!np.isEmpty()) { notes.put(j, np); total += np.size(); }
                    }
                }
                int pageNo = 0;
                if (xlsx2pdf.render.ValueDump.ON) for (SheetJob j : jobs) dumpValues(j);
                for (SheetJob j : jobs) {
                    if (j.ps.firstPageNumber != null) pageNo = j.ps.firstPageNumber - 1;
                    xlsx2pdf.render.Warnings.sheet(j.sheet.getSheetName());
                    xlsx2pdf.render.ValueDump.sheet(j.sheet.getSheetName());
                    HeaderFooter hf = new HeaderFooter(ctx, j.sheet);
                    CellPainter cp = j.chart ? null : new CellPainter(ctx, j.layout, j.ps, new CondFormat(ctx, j.sheet));
                    for (int p = 0; p < j.pages.size(); p++) {
                        checkDeadline();
                        pageNo++;
                        renderPage(ctx, doc, j, j.pages.get(p), p, pageNo, total, hf, cp);
                    }
                    List<List<String[]>> np = notes.get(j);
                    if (np != null) {
                        for (int p = 0; p < np.size(); p++) {
                            pageNo++;
                            renderCommentPage(ctx, doc, j, np.get(p), j.pages.size() + p, pageNo, total, hf);
                        }
                    }
                }
                if (doc.getNumberOfPages() == 0) {
                    // Excel は空のブックを印刷しない。有効な PDF にするため白紙を 1 ページ出す
                    doc.addPage(new PDPage(PDRectangle.A4));
                }
                PDDocumentInformation info = doc.getDocumentInformation();
                info.setProducer("xlsx2pdf");
                if (name != null) info.setTitle(name);
                sink.save(doc);
            }
        } finally {
            warnings = collecting.list();
            xlsx2pdf.render.Warnings.stop();
        }
    }

    /**
     * ToUnicode の部首を通常の漢字に直す。PDFBox はサブセットフォントの ToUnicode をフォントの cmap から作り、
     * グリフの最小のコードポイントを使う。CJK フォントは「氏」などを康煕部首 (U+2F52) からも対応付けているため、
     * コピーした文字列が部首になってしまう。対応先が U+2E80〜U+2FDF のものを NFKC の文字に書き換える。
     */
    private void finish(File pdf) throws IOException {
        try (PDDocument d = Loader.loadPDF(pdf)) {
            if (fixToUnicode(d) | protect(d)) {
                Path target = pdf.toPath().toAbsolutePath();
                Path tmp = uniqueTemp(target);
                try {
                    copyPermissions(target, tmp);
                    d.save(tmp.toFile());
                    d.close();
                    moveOver(tmp, target);
                } finally {
                    Files.deleteIfExists(tmp);
                }
            }
        }
    }

    /** メモリの PDF を仕上げる ({@link #finish(File)} と同じ)。 */
    private byte[] finish(byte[] pdf) throws IOException {
        try (PDDocument d = Loader.loadPDF(pdf)) {
            if (!(fixToUnicode(d) | protect(d))) return pdf;
            java.io.ByteArrayOutputStream o = new java.io.ByteArrayOutputStream(pdf.length);
            d.save(o);
            return o.toByteArray();
        }
    }

    /**
     * 編集の禁止を設定する ({@link Options#noEdit})。暗号化は AES-256。
     *
     * @return 設定したか
     */
    private boolean protect(PDDocument d) throws IOException {
        if (!opt.noEdit) return false;
        org.apache.pdfbox.pdmodel.encryption.AccessPermission ap = new org.apache.pdfbox.pdmodel.encryption.AccessPermission();
        ap.setCanModify(false);
        ap.setCanModifyAnnotations(false);
        ap.setCanFillInForm(false);
        ap.setCanAssembleDocument(false);
        String owner = opt.ownerPassword;
        if (owner == null || owner.isEmpty()) {
            byte[] r = new byte[24];
            new java.security.SecureRandom().nextBytes(r);
            owner = java.util.HexFormat.of().formatHex(r);
        }
        org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy p =
                new org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy(owner, "", ap);
        p.setEncryptionKeyLength(256);
        d.protect(p);
        return true;
    }

    /** 文書の ToUnicode を直す。@return 変更したか */
    private static boolean fixToUnicode(PDDocument d) throws IOException {
        boolean changed = false;
        {
            Set<COSBase> seen = new HashSet<>();
            for (PDPage pg : d.getPages()) {
                PDResources res = pg.getResources();
                if (res == null) continue;
                for (COSName fn : res.getFontNames()) {
                    COSDictionary fd = (COSDictionary) res.getCOSObject().getCOSDictionary(COSName.FONT).getDictionaryObject(fn);
                    if (fd == null || !seen.add(fd)) continue;
                    if (!(fd.getDictionaryObject(COSName.TO_UNICODE) instanceof COSStream st)) continue;
                    String text;
                    try (InputStream is = st.createInputStream()) {
                        text = new String(is.readAllBytes(), StandardCharsets.ISO_8859_1);
                    }
                    StringBuilder sb = new StringBuilder();
                    boolean any = false;
                    for (String line : text.split("\n", -1)) {
                        Matcher m = RADICAL_ENTRY.matcher(line);
                        if (m.matches()) {
                            int src = Integer.parseInt(m.group(3), 16);
                            int vf = VERTICAL_FORM_SRC.indexOf(src);
                            String rep = vf >= 0 ? String.valueOf(VERTICAL_FORM_DST.charAt(vf))
                                    : Normalizer.normalize(new String(Character.toChars(src)), Normalizer.Form.NFKC);
                            if (rep.length() == 1 && (rep.charAt(0) >= 0x3400 || vf >= 0)) {
                                line = m.group(1) + String.format("<%04X>", (int) rep.charAt(0)) + m.group(4);
                                any = true;
                            }
                        }
                        sb.append(line).append('\n');
                    }
                    sb.setLength(sb.length() - 1);
                    if (!any) continue;
                    try (OutputStream os = st.createOutputStream()) {
                        os.write(sb.toString().getBytes(StandardCharsets.ISO_8859_1));
                    }
                    changed = true;
                }
            }
        }
        return changed;
    }

    /** パッケージを開く。暗号化された (OLE2 で包まれた) ブックは復号する。 */
    /**
     * 上限を設定する: zip の展開 (POI の ZipSecureFile)、画像の画素数。上限を外すときは、POI 自身の 1 項目の読み込みの上限も外す。
     */
    private void applyLimits() {
        org.apache.poi.openxml4j.util.ZipSecureFile.setMaxEntrySize(opt.maxZipEntryBytes > 0 ? opt.maxZipEntryBytes : 0xFFFFFFFFL);
        org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(Math.max(0, opt.minInflateRatio));
        org.apache.poi.openxml4j.util.ZipSecureFile.setMaxFileCount(opt.maxZipEntries > 0 ? opt.maxZipEntries : Long.MAX_VALUE);
        xlsx2pdf.render.Images.setMaxPixels(opt.maxImagePixels);
        org.apache.poi.util.IOUtils.setByteArrayMaxOverride(opt.maxZipEntryBytes <= 0 ? Integer.MAX_VALUE - 8 : -1);
    }

    private OPCPackage openPackage(File in) throws IOException, InvalidFormatException {
        FileMagic magic;
        try (InputStream is = FileMagic.prepareToCheckMagic(new java.io.BufferedInputStream(new java.io.FileInputStream(in)))) {
            magic = FileMagic.valueOf(is);
        }
        if (magic != FileMagic.OLE2) {
            // Strict OOXML は名前空間を通常の形に書き換えてから開く
            boolean strict;
            try (InputStream is = new java.io.BufferedInputStream(new java.io.FileInputStream(in))) {
                strict = StrictOoxml.isStrict(is, opt.maxZipEntryBytes);
            } catch (IOException | RuntimeException e) {
                strict = false;
            }
            if (strict) {
                try (InputStream is = new java.io.BufferedInputStream(new java.io.FileInputStream(in))) {
                    return OPCPackage.open(new java.io.ByteArrayInputStream(StrictOoxml.toTransitional(is, opt.maxZipEntryBytes, 1L << 30)));
                }
            }
            return OPCPackage.open(in, PackageAccess.READ);
        }
        try (POIFSFileSystem fs = new POIFSFileSystem(in, true)) {
            return decrypt(fs, in.toString());
        }
    }

    /** メモリの入力からパッケージを開く。 */
    private OPCPackage openPackage(byte[] book, String name) throws IOException, InvalidFormatException {
        FileMagic magic;
        try (InputStream is = FileMagic.prepareToCheckMagic(new java.io.ByteArrayInputStream(book))) {
            magic = FileMagic.valueOf(is);
        }
        if (magic != FileMagic.OLE2) {
            byte[] conv = null;
            try {
                conv = StrictOoxml.convertIfStrict(book, opt.maxZipEntryBytes);
            } catch (RuntimeException e) {
                conv = null;
            }
            return OPCPackage.open(new java.io.ByteArrayInputStream(conv != null ? conv : book));
        }
        try (POIFSFileSystem fs = new POIFSFileSystem(new java.io.ByteArrayInputStream(book))) {
            return decrypt(fs, name);
        }
    }

    /** OLE2 で包まれたブックを復号する。 */
    private OPCPackage decrypt(POIFSFileSystem fs, String name) throws IOException, InvalidFormatException {
        if (!fs.getRoot().hasEntry("EncryptionInfo")) {
            throw new IOException("Excel 97-2003 形式 (.xls) には対応していません: " + name);
        }
        Decryptor d = Decryptor.getInstance(new EncryptionInfo(fs));
        boolean ok;
        try {
            ok = d.verifyPassword(Decryptor.DEFAULT_PASSWORD) || (opt.password != null && d.verifyPassword(opt.password));
        } catch (java.security.GeneralSecurityException e) {
            throw new IOException(e);
        }
        if (!ok) throw new IOException("パスワードで保護されたブックです (--password で指定してください): " + name);
        try (InputStream is = d.getDataStream(fs)) {
            return OPCPackage.open(is);
        } catch (java.security.GeneralSecurityException e) {
            throw new IOException(e);
        }
    }

    /** 値の書き出し (開発用): ページ数・改ページ・画面の行高と列幅 (scripts/dump-excel.ps1 と同じ形)。 */
    private static void dumpValues(SheetJob j) {
        xlsx2pdf.render.ValueDump.sheet(j.sheet.getSheetName());
        java.util.TreeSet<Integer> hb = new java.util.TreeSet<>(), vb = new java.util.TreeSet<>();
        double scale = 1;
        for (Paginator.Page p : j.pages) {
            hb.add(p.bodyR0 + 1);
            vb.add(p.bodyC0 + 1);
            scale = p.scale;
        }
        if (!hb.isEmpty()) hb.remove(hb.first());
        if (!vb.isEmpty()) vb.remove(vb.first());
        String h = hb.stream().map(String::valueOf).collect(java.util.stream.Collectors.joining(","));
        String v = vb.stream().map(String::valueOf).collect(java.util.stream.Collectors.joining(","));
        xlsx2pdf.render.ValueDump.add("page", j.pages.size(), h, v, scale * 100);
        if (j.layout == null) return;
        StringBuilder rows = new StringBuilder(), cols = new StringBuilder();
        for (int r = 0; r < 300; r++) rows.append(r == 0 ? "" : ",").append(xlsx2pdf.render.ValueDump.fmt(j.layout.screenRowPt(r)));
        for (int c = 0; c < 60; c++) cols.append(c == 0 ? "" : ",").append(xlsx2pdf.render.ValueDump.fmt(j.layout.screenColPt(c)));
        xlsx2pdf.render.ValueDump.add("rows", rows);
        xlsx2pdf.render.ValueDump.add("cols", cols);
    }

    /**
     * シートの印刷を準備する (ページ設定、印刷範囲、レイアウト、ページ分割、空ページの除去)。
     *
     * @return 準備結果。印刷するものがなければ null
     */
    private SheetJob prepare(RenderContext ctx, XSSFSheet sheet, int index) {
        // 用紙の指定がないブックは、Excel が PC のプリンタの既定の用紙を使う。既定は日本語版 Windows の PC (A4)、
        // --author-env ではブックの言語から推測する (日本語なら A4、それ以外は Letter)
        int paper = opt.defaultPaper > 0 ? opt.defaultPaper : ctx.japanese || !ctx.printer.authorEnv ? 9 : 1;
        PrintSetup ps = PrintSetup.read(ctx.wb, sheet, index, paper, opt.printerDpi);
        if (!ctx.printer.authorEnv) ps.minMargin = 0;
        if (sheet instanceof XSSFChartSheet) {
            DrawingPainter dp = new DrawingPainter(ctx, sheet, null);
            List<Paginator.Page> pages = new ArrayList<>();
            Paginator.Page pg = new Paginator.Page();
            pg.scale = 1;
            pg.offsetX = ps.left;
            pg.offsetY = ps.top;
            pg.contentW = ps.paperW - ps.left - ps.right;
            pg.contentH = ps.paperH - ps.top - ps.bottom;
            pages.add(pg);
            return new SheetJob(sheet, index, ps, null, pages, dp, true);
        }
        // 使用範囲 {先頭行, 先頭列, 最終行, 最終列} を図形の範囲まで広げる
        DrawingPainter probe = new DrawingPainter(ctx, sheet, null);
        int[] used = SheetLayout.usedRange(sheet, true);
        int[] dext = probe.extent(); // {最終行, 最終列, 先頭行, 先頭列} または null
        if (dext != null) {
            // セルに内容がなく図形だけのシートは、Excel は A1 から印刷する (図形の上・左の空きも印刷する)
            if (used == null) used = new int[] {0, 0, dext[0], dext[1]};
            else {
                used[0] = Math.min(used[0], dext[2]);
                used[1] = Math.min(used[1], dext[3]);
                used[2] = Math.max(used[2], dext[0]);
                used[3] = Math.max(used[3], dext[1]);
            }
        }
        // スパークラインのセルも使用範囲に含める (値のないセルに描かれることが多い)
        xlsx2pdf.render.Sparklines.Cells sparks = xlsx2pdf.render.Sparklines.cellsOf(ctx, sheet);
        for (int[] rc : sparks.list()) {
            if (used == null) used = new int[] {rc[0], rc[1], rc[0], rc[1]};
            else {
                used[0] = Math.min(used[0], rc[0]); used[1] = Math.min(used[1], rc[1]);
                used[2] = Math.max(used[2], rc[0]); used[3] = Math.max(used[3], rc[1]);
            }
        }
        List<CellRangeAddress> areas = new ArrayList<>();
        if (!ps.printAreas.isEmpty()) {
            for (CellRangeAddress a : ps.printAreas) {
                int lr = a.getLastRow(), lc = a.getLastColumn();
                // 行全体・列全体の印刷範囲は使用範囲で終える
                if (used != null) {
                    if (lr >= 1048575) lr = Math.max(a.getFirstRow(), used[2]);
                    if (lc >= 16383) lc = Math.max(a.getFirstColumn(), used[3]);
                } else {
                    if (lr >= 1048575) lr = a.getFirstRow();
                    if (lc >= 16383) lc = a.getFirstColumn();
                }
                areas.add(new CellRangeAddress(a.getFirstRow(), lr, a.getFirstColumn(), lc));
            }
        } else {
            if (used == null) return null;
            areas.add(new CellRangeAddress(used[0], used[2], used[1], used[3]));
        }
        // レイアウトを計算する行数・列数
        int needR = 0, needC = 0;
        for (CellRangeAddress a : areas) {
            needR = Math.max(needR, a.getLastRow() + 1);
            needC = Math.max(needC, a.getLastColumn() + 1);
        }
        if (ps.titleRows != null) needR = Math.max(needR, ps.titleRows.getLastRow() + 1);
        if (ps.titleCols != null) needC = Math.max(needC, ps.titleCols.getLastColumn() + 1);
        if (dext != null) {
            needR = Math.max(needR, dext[0] + 2);
            needC = Math.max(needC, dext[1] + 2);
        }
        boolean ovRange = ps.printAreas.isEmpty();
        if (ovRange) needC += OV_RANGE_COLS;   // 文字のはみ出しで広がる分
        SheetLayout layout = new SheetLayout(sheet, ctx.mdw, ctx.printer, ps.hdpi, ps.vdpi, needR + 1, needC + 1, opt.screenDpi);
        // 数式の表示のシートは、文字のはみ出しで範囲を広げない (exp_showf)
        if (ovRange && !layout.showFormulas) {
            CellRangeAddress a = areas.get(0);
            int lc = new CellPainter(ctx, layout, ps, new CondFormat(ctx, sheet)).overflowLastColumn(a.getLastRow(), a.getLastColumn(), a.getLastColumn() + OV_RANGE_COLS);
            if (lc > a.getLastColumn()) areas.set(0, new CellRangeAddress(a.getFirstRow(), a.getLastRow(), a.getFirstColumn(), lc));
        }
        Paginator paginator = new Paginator(layout, ps);
        if (ps.headings && ctx.defaultFontName != null)
            paginator.digitAdvance = ctx.fonts.spec(ctx.defaultFontName, false, false).charWidth('0', (float) ctx.defaultFontSize, ctx.fonts);
        if (opt.maxPages > 0 && pagesMade >= opt.maxPages) {
            throw new LimitExceededException("ページ数が上限 (" + opt.maxPages + ") を超えます (シート「" + sheet.getSheetName() + "」)。--max-pages で上限を変更できます", false);
        }
        paginator.setMaxPages(opt.maxPages > 0 ? opt.maxPages - pagesMade : 0);
        checkDeadline();
        List<Paginator.Page> pages = paginator.paginate(areas);
        pagesMade += pages.size();
        checkDeadline();
        if (pages.isEmpty()) return null;
        DrawingPainter dp = new DrawingPainter(ctx, sheet, layout);
        dp.patternLeft = ps.left;
        dp.patternTop = ps.top;
        // 今の Excel は印刷範囲を指定したシートでは、何もないページも印刷する
        boolean keepArea = !ps.printAreas.isEmpty() && (!ctx.printer.authorEnv);
        if (pages.size() > 1 && !keepArea) {
            // Excel は何もないページ (本体に値・塗りつぶし・罫線・オブジェクトがない) を印刷しない
            List<Paginator.Page> kept = new ArrayList<>();
            CellPainter spill = new CellPainter(ctx, layout, ps, new CondFormat(ctx, sheet));
            boolean[] full = new boolean[pages.size()];
            for (int i = 0; i < pages.size(); i++) {
                Paginator.Page p = pages.get(i);
                layout.setScale(p.scale);
                checkDeadline();
                full[i] = !blank(sheet, layout, p, dp, spill) || sparks.any(p.segments);
            }
            for (int i = 0; i < pages.size(); i++) {
                if (full[i] || laterFull(full, i)) kept.add(pages.get(i));
            }
            if (!kept.isEmpty()) pages = kept;
        }
        return new SheetJob(sheet, index, ps, layout, pages, dp, false);
    }

    /**
     * 白紙のページは、シートの最後 (印刷の順) に続くものだけを省く。後ろに中身のあるページがある白紙は印刷する
     * (帯ごとに末尾を省く形より合う)。
     */
    private static boolean laterFull(boolean[] full, int i) {
        for (int j = i + 1; j < full.length; j++) if (full[j]) return true;
        return false;
    }

    /**
     * ページが空か判定する。本体のどのセルにも値・塗りつぶし・罫線がなく、重なるオブジェクトもなければ空。
     * 先頭行・先頭列の上・左の罫線は、前のページから続く辺なら数えない。
     */
    private static boolean blank(XSSFSheet sheet, SheetLayout L, Paginator.Page p, DrawingPainter dp, CellPainter spill) {
        Paginator.Segment b = null;
        for (Paginator.Segment s : p.segments) if (s.body) b = s;
        if (b == null) return false;
        int ac0 = b.area == null ? 0 : b.area.getFirstColumn();
        for (int r = b.r0; r <= b.r1; r++) {
            if (L.rh(r) <= 0) continue;
            XSSFRow row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = b.c0; c <= b.c1; c++) {
                if (L.cw(c) <= 0) continue;
                XSSFCell cell = row.getCell(c);
                if (cell == null) continue;
                if (!CellPainter.isEmptyCell(cell)) return false;
                XSSFCellStyle st = cell.getCellStyle();
                if (st == null) continue;
                if (st.getFillPattern() != FillPatternType.NO_FILL || CellPainter.hasGradient(cell)) return false;
                if (st.getBorderBottom() != BorderStyle.NONE || st.getBorderRight() != BorderStyle.NONE) return false;
                if (st.getBorderTop() != BorderStyle.NONE) return false;
                if (st.getBorderLeft() != BorderStyle.NONE && (c > b.c0 || c == ac0)) return false;
            }
        }
        // 結合範囲 (値や罫線がなくても) の左上のセルがあるページは白紙としない
        for (CellRangeAddress m : sheet.getMergedRegions()) {
            int mr = m.getFirstRow(), mc = m.getFirstColumn();
            if (mr >= b.r0 && mr <= b.r1 && mc >= b.c0 && mc <= b.c1 && L.rh(mr) > 0 && L.cw(mc) > 0) return false;
            // 値のある結合範囲が前のページから続くなら、文字の続きが載るので白紙としない
            if (m.getLastRow() >= b.r0 && mr <= b.r1 && m.getLastColumn() >= b.c0 && mc <= b.c1) {
                XSSFRow tr = sheet.getRow(mr);
                XSSFCell tc = tr == null ? null : tr.getCell(mc);
                if (tc != null && !CellPainter.isEmptyCell(tc) && (mc < b.c0 || mergedTextReaches(L, m, tc, b))) return false;
            }
        }
        boolean o = dp.hasObjectIn(b.r0, b.c0, b.r1, b.c1);
        if (!o && (spill.hasSpill(b) || spill.hasSlantSpill(b))) return false;
        return !o;
    }

    /**
     * 結合範囲の文字 (配置と行数の概算) が、ページの行の範囲に届くか。折り返しは文字幅を全角 1em・半角 0.5em として行数を数え、
     * 行の高さは 1.2em とする。
     */
    private static boolean mergedTextReaches(SheetLayout L, CellRangeAddress m, XSSFCell tc, Paginator.Segment b) {
        String t = tc.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING ? tc.getStringCellValue() : "0";
        XSSFCellStyle st = tc.getCellStyle();
        double size = st == null ? 11 : st.getFont().getFontHeightInPoints();
        double w = Math.max(1, L.x(m.getLastColumn() + 1) - L.x(m.getFirstColumn()) - 3);
        boolean wrap = st != null && st.getWrapText();
        int lines = 0;
        double maxW = 0;
        for (String para : t.split("\n", -1)) {
            double em = 0;
            for (int i = 0; i < para.length(); i++) em += para.charAt(i) < 0x100 || (para.charAt(i) >= 0xff61 && para.charAt(i) <= 0xff9f) ? 0.5 : 1;
            lines += wrap ? Math.max(1, (int) Math.ceil(em * size / w)) : 1;
            maxW = Math.max(maxW, Math.min(em * size, w));
        }
        // 横: 配置で文字の左右を決める (標準は文字列なら左、数値なら右とする)
        double x0 = L.x(m.getFirstColumn()), x1 = L.x(m.getLastColumn() + 1), la, lb;
        org.apache.poi.ss.usermodel.HorizontalAlignment ha = st == null ? org.apache.poi.ss.usermodel.HorizontalAlignment.GENERAL : st.getAlignment();
        switch (ha) {
            case CENTER, CENTER_SELECTION -> { la = (x0 + x1 - maxW) / 2; lb = la + maxW; }
            case RIGHT -> { lb = x1 - 1.5; la = lb - maxW; }
            case FILL, JUSTIFY, DISTRIBUTED -> { la = x0; lb = x1; }
            case GENERAL -> {
                if (tc.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING) { la = x0 + 1.5; lb = la + maxW; }
                else { lb = x1 - 1.5; la = lb - maxW; }
            }
            default -> { la = x0 + 1.5; lb = la + maxW; }
        }
        if (lb <= L.x(b.c0) + 0.5 || la >= L.x(b.c1 + 1) - 0.5) return false;
        double th = lines * size * 1.2;
        double y0 = L.y(m.getFirstRow()), y1 = L.y(m.getLastRow() + 1);
        org.apache.poi.ss.usermodel.VerticalAlignment va = st == null ? org.apache.poi.ss.usermodel.VerticalAlignment.BOTTOM : st.getVerticalAlignment();
        double ta, tb;
        switch (va) {
            case TOP -> { ta = y0; tb = y0 + th; }
            case CENTER -> { ta = (y0 + y1 - th) / 2; tb = ta + th; }
            case JUSTIFY, DISTRIBUTED -> { ta = y0; tb = y1; }
            default -> { tb = y1; ta = y1 - th; }
        }
        return tb > L.y(b.r0) + 0.5 && ta < L.y(b.r1 + 1) - 0.5;
    }

    /** 末尾のコメントの一覧の見出し (Excel の日本語版の印刷)。 */
    static final String NOTE_CELL = "セル:", NOTE_TEXT = "コメント:";

    /**
     * 末尾に印刷するコメントの一覧を、ページごとの行 ({見出し, 文字}) に分ける。コメントは行、列の順。
     * 1 件は「セル: 番地」「コメント: 文字 (改行ごとに 1 行)」と空行。
     */
    private static List<List<String[]>> commentPages(RenderContext ctx, SheetJob j) {
        List<List<String[]>> pages = new ArrayList<>();
        java.util.Map<org.apache.poi.ss.util.CellAddress, org.apache.poi.xssf.usermodel.XSSFComment> cm;
        try {
            cm = j.sheet.getCellComments();
        } catch (RuntimeException e) {
            xlsx2pdf.render.Warnings.add("コメントを読めませんでした", e);
            return pages;
        }
        if (cm == null || cm.isEmpty()) return pages;
        List<org.apache.poi.ss.util.CellAddress> keys = new ArrayList<>(cm.keySet());
        keys.sort((a, b) -> a.getRow() != b.getRow() ? Integer.compare(a.getRow(), b.getRow()) : Integer.compare(a.getColumn(), b.getColumn()));
        List<String[]> lines = new ArrayList<>();
        for (org.apache.poi.ss.util.CellAddress k : keys) {
            var c = cm.get(k);
            String t = c.getString() == null ? "" : c.getString().getString();
            if (!lines.isEmpty()) lines.add(new String[] {"", ""});
            lines.add(new String[] {NOTE_CELL, k.formatAsString()});
            String[] ls = t.replace("\r\n", "\n").replace('\r', '\n').split("\n", -1);
            for (int i = 0; i < ls.length; i++) lines.add(new String[] {i == 0 ? NOTE_TEXT : "", ls[i]});
        }
        double pitch = notePitch(ctx, j) * noteScale(j);
        int per = Math.max(1, (int) Math.floor((j.ps.paperH - j.ps.top - j.ps.bottom) / pitch));
        for (int i = 0; i < lines.size(); i += per) pages.add(lines.subList(i, Math.min(lines.size(), i + per)));
        return pages;
    }

    /** コメントの一覧の倍率: Excel はシートの印刷の倍率 (拡大縮小・1 ページに合わせる) で描く。 */
    private static double noteScale(SheetJob j) {
        return j.pages.isEmpty() || !(j.pages.get(0).scale > 0) ? 1 : j.pages.get(0).scale;
    }

    /** コメントの一覧の行送り (標準フォントの行の高さ)。 */
    private static double notePitch(RenderContext ctx, SheetJob j) {
        double p = j.layout.linePitch(ctx.defaultFontName, ctx.defaultFontSize);
        return p > 0 ? p : ctx.defaultFontSize * 1.68;
    }

    /** コメントの一覧の 1 ページを描く (ヘッダー・フッターも描く)。 */
    private void renderCommentPage(RenderContext ctx, PDDocument doc, SheetJob j, List<String[]> lines, int pageInSheet, int pageNo,
                                   int total, HeaderFooter hf) throws IOException {
        PrintSetup ps = j.ps;
        boolean snap = !ctx.printer.authorEnv;
        double mediaW = snap ? Math.rint(ps.paperW * 600 / 72) * 0.12 : ps.paperW, mediaH = snap ? Math.rint(ps.paperH * 600 / 72) * 0.12 : ps.paperH;
        PDPage page = new PDPage(new PDRectangle((float) mediaW, (float) mediaH));
        doc.addPage(page);
        try (PdfCanvas cv = new PdfCanvas(doc, page, ctx.fonts)) {
            cv.save();
            double k = noteScale(j);
            hf.paint(cv, ps, ps.paperW, ps.paperH, k, pageInSheet, pageNo, total, j.sheet.getSheetName());
            cv.restore();
            xlsx2pdf.render.FontManager.Spec sp = ctx.fonts.spec(ctx.defaultFontName, false, false);
            float size = (float) (ctx.defaultFontSize * k);
            double pitch = notePitch(ctx, j) * k;
            double labelW = sp.width(NOTE_TEXT, size, ctx.fonts);
            double x0 = ps.left, xv = x0 + labelW + NOTE_GAP * k;
            // 最初の基線: 上余白 + フォントの大きさ × 0.993 (Excel の印刷: 游ゴシック 11pt で 10.92pt)
            double y = ps.top + size * 0.993;
            for (String[] l : lines) {
                if (!l[0].isEmpty()) cv.text(l[0], xv - NOTE_GAP * k - sp.width(l[0], size, ctx.fonts), y, sp, size, java.awt.Color.BLACK, 1, 0);
                if (!l[1].isEmpty()) cv.text(l[1], xv, y, sp, size, java.awt.Color.BLACK, 1, 0);
                y += pitch;
            }
        }
    }

    /** グラフシートのグラフは印刷範囲より幅 0.1in・高さ 9pt 小さい (左上をそろえる) */
    static final double CS_DW = 7.2, CS_DH = 9;

    /** コメントの一覧の見出しと文字の間 (pt)。 */
    static final double NOTE_GAP = 3.43;

    /**
     * 1 ページを描く (セル、図形、ヘッダー/フッター)。
     *
     * @param pageInSheet シート内のページ番号 (0 始まり)
     * @param pageNo      印刷するページ番号
     * @param total       総ページ数
     */
    private void renderPage(RenderContext ctx, PDDocument doc, SheetJob j, Paginator.Page pg, int pageInSheet, int pageNo,
                            int total, HeaderFooter hf, CellPainter cp) throws IOException {
        PrintSetup ps = j.ps;
        double pageScale = opt.excelPageSize ? pg.scale : 1;
        // 用紙はプリンタの画素に丸めた大きさ (Excel の A4 は 595.32 × 841.92pt = 4961 × 7016 画素)
        boolean snap = !ctx.printer.authorEnv;
        double mediaW = snap ? Math.rint(ps.paperW * 600 / 72) * 0.12 : ps.paperW, mediaH = snap ? Math.rint(ps.paperH * 600 / 72) * 0.12 : ps.paperH;
        PDPage page = new PDPage(new PDRectangle((float) (mediaW / pageScale), (float) (mediaH / pageScale)));
        doc.addPage(page);
        try (PdfCanvas cv = new PdfCanvas(doc, page, ctx.fonts)) {
            if (pageScale != 1) cv.scale(1 / pageScale, 1 / pageScale);
            // Microsoft Print to PDF は画素を 用紙の pt ÷ (画素数 − 1) で変換する (A4: 横 4960 画素、縦 7016 画素。全倍率の罫線の実測)
            if (snap) {
                double nx = Math.rint(mediaW / 0.12), ny = Math.rint(mediaH / 0.12);
                cv.scale(nx / (nx - 1), ny / (ny - 1));
            }
            // ヘッダー・フッターは本文より先に描く (Excel と同じく、塗ったセルが重なればヘッダーが隠れる)
            cv.save();
            hf.paint(cv, ps, ps.paperW, ps.paperH, pg.scale, pageInSheet, pageNo, total, j.sheet.getSheetName());
            cv.restore();
            if (j.chart) {
                cv.save();
                j.drawings.paintChartSheet(cv, pg.offsetX, pg.offsetY, pg.contentW - CS_DW, pg.contentH - CS_DH);
                cv.restore();
            } else {
                // 右から左のシート: 本体を用紙の中心で左右反転して描く (A 列が右端。文字と図形の中身は反転を打ち消す)
                j.drawings.bw = ps.blackAndWhite;
                if (ps.rtl) {
                    cv.save();
                    cv.transform(PdfCanvas.flipAbout(ps.paperW / 2));
                    j.drawings.rtl = true;
                }
                for (Paginator.Segment s : pg.segments) {
                    cv.save();
                    cv.translate(pg.offsetX + s.dx * pg.scale, pg.offsetY + s.dy * pg.scale);
                    cv.scale(pg.scale, pg.scale);
                    double sx = j.layout.x(s.c0), sy = j.layout.y(s.r0);
                    cv.translate(-sx, -sy);
                    double w = j.layout.x(s.c1 + 1) - sx, h = j.layout.y(s.r1 + 1) - sy;
                    if (pg.hdrW > 0 || pg.hdrH > 0) {
                        // 行列番号 (ブロックの外側に描くので、切り抜きの前に描く)
                        boolean top = true, left = true, right = true, bottom = true;
                        for (Paginator.Segment o : pg.segments) {
                            if (o.dy < s.dy) top = false;
                            if (o.dx < s.dx) left = false;
                            if (o.dy > s.dy) bottom = false;
                            if (o.dx > s.dx) right = false;
                        }
                        cv.mirrorText = ps.rtl;
                        cp.paintHeadings(cv, s, pg.hdrW, pg.hdrH, top, left, right, bottom);
                        cv.mirrorText = false;
                    }
                    // Excel のクリップはブロックの外側に約 2pt (用紙上) 広がり、端の罫線が半分に切れない
                    double e = 2.0 / pg.scale;
                    cv.clipRect(sx - e, sy - e, w + 2 * e, h + 2 * e);
                    cp.setBlockPlacement(pg.offsetX + s.dx * pg.scale, pg.offsetY + s.dy * pg.scale, sx, sy);
                    cv.mirrorText = ps.rtl;
                    cp.paint(cv, s);
                    cv.mirrorText = false;
                    if (s.body && !ps.draft) {
                        // 図形・グラフの切り抜きは外側に約 0.44pt (用紙上。ページの右端で切れる位置)
                        double ed = DRAWING_CLIP / pg.scale;
                        cv.clipRect(sx - ed, sy - ed, w + 2 * ed, h + 2 * ed);
                        j.drawings.paint(cv, sx, sy, sx + w, sy + h); // 簡易印刷は図形・グラフを印刷しない
                    }
                    cv.restore();
                }
                if (ps.rtl) cv.restore();
            }
        }
    }
}
