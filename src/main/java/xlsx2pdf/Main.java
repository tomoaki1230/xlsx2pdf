package xlsx2pdf;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * コマンドラインの入口。
 *
 * <pre>
 * java -jar xlsx2pdf.jar input.xlsx output.pdf [--font-dir DIR]... [--sheet NAME]... [--active-only] [--excel-page-size]
 *      [--default-paper A4] [--screen-dpi 96] [--printer-dpi 600]
 *      [--max-pages N] [--max-image-pixels N] [--max-zip-entry-mb N] [--timeout SEC] [--no-limits] [--verbose]
 * </pre>
 *
 * 終了コード: 0 成功、1 変換できない (形式・パスワード・想定外のエラー)、2 引数の誤り、3 上限超過・メモリ不足、4 時間切れ。
 */
public final class Main {
    private Main() {
    }

    /**
     * コマンドラインのオプションを解析する。
     *
     * @param args  コマンドライン引数
     * @param files オプション以外の引数 (入力・出力ファイル) を追加するリスト
     * @return 変換オプション。{@code --help} が指定されたときは {@code null}
     */
    public static Converter.Options parseOptions(String[] args, List<String> files) {
        Converter.Options opt = new Converter.Options();
        boolean noLimits = false;
        for (int i = 0; i < args.length; i++) {
            switch (args[i]) {
                case "--font-dir" -> opt.fontDirs.add(new File(args[++i]));
                case "--sheet" -> opt.sheets.add(args[++i]);
                case "--active-only" -> opt.activeOnly = true;
                case "--excel-page-size" -> opt.excelPageSize = true;
                case "--default-paper" -> opt.defaultPaper = xlsx2pdf.render.PrintSetup.paperCode(args[++i]);
                case "--password" -> opt.password = args[++i];
                case "--use-file-dpi" -> xlsx2pdf.render.PrintSetup.useFileDpi = true;
                case "--screen-dpi" -> opt.screenDpi = dpiArg(args[++i], "--screen-dpi");
                case "--printer-dpi" -> opt.printerDpi = dpiArg(args[++i], "--printer-dpi");
                case "--author-env" -> opt.authorEnv = true;
                case "--max-pages" -> opt.maxPages = longArg(args[++i], "--max-pages");
                case "--max-image-pixels" -> opt.maxImagePixels = longArg(args[++i], "--max-image-pixels");
                case "--max-zip-entry-mb" -> opt.maxZipEntryBytes = longArg(args[++i], "--max-zip-entry-mb") << 20;
                case "--timeout" -> opt.timeoutSeconds = longArg(args[++i], "--timeout");
                case "--no-limits" -> noLimits = true;
                case "--verbose" -> verbose = true;
                case "-h", "--help" -> { return null; }
                default -> {
                    if (args[i].startsWith("--")) throw new IllegalArgumentException("不明なオプションです: " + args[i]);
                    files.add(args[i]);
                }
            }
        }
        // --no-limits はほかの上限の指定より優先する (順序によらない)
        if (noLimits) opt.noLimits();
        return opt;
    }

    /** POI の zip の上限 (ZipSecureFile の zip 爆弾の検出・項目の大きさ) によるエラーか (原因をたどって判定)。 */
    static boolean zipLimit(Throwable e) {
        for (Throwable t = e; t != null; t = t.getCause()) {
            String m = String.valueOf(t.getMessage());
            if (m.contains("Zip bomb") || m.contains("Max entry size") || m.contains("exceeds the max") || m.contains("too large")) return true;
            if (t.getCause() == t) break;
        }
        return false;
    }

    /** xlsx として読めない入力によるエラーか。 */
    static boolean notXlsx(Throwable e) {
        for (Throwable t = e; t != null; t = t.getCause()) {
            if (t instanceof org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException
                    || t instanceof org.apache.poi.poifs.filesystem.NotOLE2FileException
                    || t instanceof org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException
                    || t instanceof java.util.zip.ZipException) return true;
            if (t.getCause() == t) break;
        }
        return false;
    }

    /** 想定外のエラーの詳細 (スタックトレース) を出すか (--verbose)。 */
    private static boolean verbose;

    /** 0 以上の整数の引数を読む (0 = 無制限)。 */
    static long longArg(String v, String opt) {
        try {
            long n = Long.parseLong(v.trim().replace("_", "").replace(",", ""));
            if (n < 0) throw new NumberFormatException();
            return n;
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException(opt + " の値が不正です (0 以上の整数。0 = 無制限): " + v);
        }
    }

    /**
     * DPI の引数を読む。「100%」「150%」のような画面の拡大率も受け付ける (96 dpi × 拡大率)。
     *
     * @throws IllegalArgumentException 解釈できない値・範囲外 (72〜4800) のとき
     */
    static int dpiArg(String v, String opt) {
        String s = v.trim();
        int dpi;
        try {
            dpi = s.endsWith("%") ? (int) Math.round(96 * Double.parseDouble(s.substring(0, s.length() - 1)) / 100) : Integer.parseInt(s);
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException(opt + " の値が不正です: " + v);
        }
        if (dpi < 72 || dpi > 4800) throw new IllegalArgumentException(opt + " の値が範囲外です (72〜4800): " + v);
        return dpi;
    }

    /**
     * 入力 xlsx を PDF に変換する。
     *
     * @param args {@code 入力.xlsx 出力.pdf [オプション]}
     * @throws Exception 変換以外の想定外のエラー
     */
    public static void main(String[] args) throws Exception {
        System.setProperty("java.awt.headless", "true");
        List<String> files = new ArrayList<>();
        Converter.Options opt;
        try {
            opt = parseOptions(args, files);
        } catch (IllegalArgumentException | ArrayIndexOutOfBoundsException e) {
            System.err.println(e instanceof ArrayIndexOutOfBoundsException ? "オプションの値がありません" : e.getMessage());
            usage();
            System.exit(2);
            return;
        }
        if (opt == null) { usage(); return; }
        if (files.size() != 2) {
            usage();
            System.exit(2);
        }
        File in = new File(files.get(0));
        File out = new File(files.get(1));
        if (!in.isFile()) {
            System.err.println("input not found: " + in);
            System.exit(1);
        }
        if (Converter.samePath(in, out)) {
            System.err.println("入力と出力が同じファイルです: " + out);
            System.exit(2);
        }
        boolean existed = out.exists();
        int code = 0;
        try {
            new Converter(opt).convert(in, out);
        } catch (LimitExceededException e) {
            System.err.println("変換を中止しました: " + e.getMessage());
            code = e.isTimeout() ? 4 : 3;
        } catch (FontUnavailableException e) {
            System.err.println("変換できませんでした: " + e.getMessage());
            code = 1;
        } catch (OutOfMemoryError e) {
            System.err.println("メモリが足りないため変換できませんでした。java の -Xmx で使えるメモリを増やしてください"
                    + " (上限の既定値は --max-pages などで小さくすることもできます)");
            if (verbose) e.printStackTrace();
            code = 3;
        } catch (java.io.IOException | org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            System.err.println("変換できませんでした: " + e.getMessage());
            code = 1;
        } catch (Throwable e) {
            if (zipLimit(e)) {
                System.err.println("変換を中止しました: xlsx の中に、極端に圧縮された・大きすぎる項目があります (zip 爆弾の疑い)。"
                        + "--max-zip-entry-mb で上限を上げるか、--no-limits で外せます");
                code = 3;
            } else if (notXlsx(e)) {
                System.err.println("変換できませんでした: xlsx (Office Open XML) として読めないファイルです");
                code = 1;
            } else {
                // 想定外のエラー: 内部の詳細 (スタックトレース) は --verbose のときだけ出す
                System.err.println("変換できませんでした (想定外のエラー: " + e.getClass().getSimpleName() + ")。詳細は --verbose で表示できます");
                code = 1;
            }
            if (verbose) e.printStackTrace();
        }
        if (code != 0) {
            // 途中まで書いた出力は残さない (もとからあったファイルは消さない)
            if (!existed && out.exists() && !out.delete()) out.deleteOnExit();
            System.exit(code);
        }
    }

    /** 使い方を標準エラーに出力する。 */
    private static void usage() {
        System.err.println("""
                usage: java -jar xlsx2pdf.jar <input.xlsx> <output.pdf> [options]
                  --font-dir DIR      additional folder with TrueType fonts (e.g. a copy of Windows fonts)
                  --sheet NAME        print only this sheet (repeatable)
                  --active-only       print only selected sheets (Excel's default "active sheets")
                  --excel-page-size   page size = paper size / print scale, exactly like Excel's PDF export
                  --default-paper P   paper used when the file has no paper size (A4, A3, B4, B5, Letter, Legal...)
                  --password PW       password of an encrypted workbook
                  --use-file-dpi      lay out with the printer resolution stored in the file (default: 600 dpi)
                  --screen-dpi N      display DPI of the PC whose Excel output you want to match (96 = 100%, 120 = 125%,
                                      144 = 150%, 192 = 200%; "150%" is also accepted). Default: 96
                  --author-env        match the PDF the author exported: infer the author's Excel version and display
                                      DPI from the file (old Excel, high-DPI screen). Default: current Excel on a 100% PC
                  --printer-dpi N     resolution of the printer selected in that Excel (Microsoft Print to PDF = 600,
                                      many laser printers 300 or 600). Default: 600
                limits (protection against malicious files; 0 = unlimited):
                  --max-pages N         pages laid out before blank pages are dropped (default 10000)
                  --max-image-pixels N  pixels of one embedded image (default 50000000)
                  --max-zip-entry-mb N  uncompressed size of one part of the xlsx (default 1024)
                  --timeout SEC         conversion time limit in seconds (default 0 = unlimited)
                  --no-limits           remove all limits (for large files when memory and time are plentiful;
                                        give java enough memory, e.g. java -Xmx16g -jar ...)
                  --verbose             print details of unexpected errors
                exit codes: 0 ok, 1 cannot convert, 2 bad arguments, 3 limit exceeded / out of memory, 4 timeout
                """);
    }
}
