package xlsx2pdf.render;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;

/** 印刷範囲を Excel と同じ規則でページに分割する。 */
public final class Paginator {
    /** ページ上の位置に描くセルの矩形ブロック (座標はシート単位の pt、拡大縮小前)。 */
    public static final class Segment {
        /** 行・列の範囲 (両端を含む、0 始まり)。 */
        public final int r0, r1, c0, c1;
        /** 内容領域の左上からのオフセット (拡大縮小前の pt)。 */
        public final double dx, dy;
        /** 本体 (印刷タイトルでない) ブロックか。 */
        public final boolean body;
        /** ページが属する印刷範囲。 */
        public CellRangeAddress area;
        /** 文字だけに足す位置の補正 (pt。POSM_LINE100)。 */
        public double textDx, textDy;

        Segment(int r0, int r1, int c0, int c1, double dx, double dy, boolean body) {
            this.r0 = r0; this.r1 = r1; this.c0 = c0; this.c1 = c1; this.dx = dx; this.dy = dy; this.body = body;
        }
    }

    /** 1 ページ。 */
    public static final class Page {
        /** 描くブロック (印刷タイトルと本体)。 */
        public final List<Segment> segments = new ArrayList<>();
        /** 拡大縮小率 (1.0 = 100%)。 */
        public double scale;
        /** 内容の左上の位置 (用紙上の pt)。 */
        public double offsetX, offsetY;
        /** 内容の大きさ (拡大縮小前の pt)。 */
        public double contentW, contentH;
        /** 行列番号の印刷: 行番号の列の幅と列の見出しの高さ (シートの pt。印刷しなければ 0)。 */
        public double hdrW, hdrH;
        /** ページが属する印刷範囲全体。 */
        public CellRangeAddress area;
        /** 本体の行・列の範囲 (両端を含む)。 */
        public int bodyR0, bodyR1, bodyC0, bodyC1;
    }

    private final SheetLayout L;
    private final PrintSetup ps;
    /** 作るページ数 (白紙を除く前) の上限 (0 以下は無制限) と、これまでに作った数。 */
    private long maxPages, made;

    /** 作るページ数の上限を設定する (0 以下は無制限。超えると {@link xlsx2pdf.LimitExceededException})。 */
    public void setMaxPages(long maxPages) {
        this.maxPages = maxPages;
    }

    /** ページ分割器を作る。 */
    public Paginator(SheetLayout layout, PrintSetup ps) {
        this.L = layout;
        this.ps = ps;
    }

    /** ページに入る行または列の範囲 (両端を含む)。 */
    private record Band(int a, int b) {}

    /** 縮小したシートの横の補正。 */
    static final double SCALED_DX = 0.8;
    /** 縮小したシートの縦の補正。 */
    static final double SCALED_DY = 0.7;

    /** 中央に印刷するときの表の中心の、使える範囲の中央からのずれ (pt、上・左へ): 4 画素、細い線の太さが奇数なら 3.5 画素。 */
    private static double posxCenter(double scale) {
        long w = (long) Math.floor(8 * scale + 0.5 + 1e-9);
        return (w % 2 == 1 ? 3.5 : 4) * 0.12;
    }

    /** 罫線の中心のセルの境界からのずれ (pt): 4 画素、細い線の太さ round(8 × 倍率) が奇数なら 4.5 画素。 */
    private static double posxOffset(double scale) {
        long w = (long) Math.floor(8 * scale + 0.5 + 1e-9);
        return (w % 2 == 1 ? 4.5 : 4) * 0.12;
    }
    static final double POSM_DX = 0.5;
    /** 既定の動きの 100% のシートの表の位置の補正 (pt)。列幅の丸め の後に掃引し直した値。 */
    static final double POSM_DX100_MODERN = 0.75;
    /** 倍率 100% の表の縦の補正。 */
    static final double POSM_DY100_MODERN = 0.75;
    static final double POSM_CENTER_DX_MODERN = -0.2;
    /**
     * 線形に縮小する許容幅 (%): 新しいブック (lastEdited ≥ 7 か不明) は、保存された倍率が線形の計算値より 1 小さくてもよい。
     * 収まり具合がきわどいと、作成した PC の計算は 1% 小さくなる。
     * 許容幅 0: x21 0.7552・webnew 0.8477・webold 0.9285、1 (新しいブックだけ): 0.7570・0.8507・0.9285、1 (すべて): 0.7586・0.8492・0.9281。
     */
    static final int FITLIN_TOL = 1;
    static final double FIT_TIGHT_DEFAULT = 0.0012;
    /**
     * 既定の動きの列の改ページ: 列の合計が使える幅よりこの値 (pt) だけ大きくても収まるとみなす。負なら小さくなければ収まらない。
     * --author-env は 0.01。
     */
    static final double COL_TOL = -0.05;
    /**
     * 縮小したシートの改ページで、行の合計が使える高さより 0.05pt 以上小さくなければ収まらないとみなす。列は 0 (効果 +0.0001 で、ちょうど同じ幅のシートを Excel が収める例と収めない例がある)。
     */
    static final double ROW_STRICT = 0.05;
    /** 「ページに合わせる」の倍率を求めるときは、列・行を縮小後の画素に丸めない。 */
    /**
     * 「ページに合わせる」の倍率の判定の余裕 (プリンタの画素。既定の動きのみ)。列・行は倍率ごとに四捨五入して足し、整数の使える大きさ ({@link #intAvail}) より
     * 8 画素を超えて小さければ収まる (余り 8 以下は収まらず 9 以上は収まる)。
     */
    static final double FIT_PXW = 8.5;
    static final double FIT_PXH = 8.5;
    /**
     * 行の改ページで、整数の使える高さ ({@link #intAvail}) から引くプリンタの画素 (既定の動きのみ)。用紙によらず一定。
     */
    static final double PAGE_PXH = 8.0;

    /**
     * 印刷範囲をページに分割する。拡大縮小率を決め、レイアウトにも設定する。
     *
     * <p>「ページに合わせる」の倍率:
     * <ol>
     *   <li>保存された倍率が、100% のレイアウトをそのまま縮小したときに収まる最大の倍率 (線形の計算値) と一致するなら、Excel はその倍率で
     *       線形に縮小している (フォントを倍率で実体化し直さない)。</li>
     *   <li>そうでなければ、その倍率で実体化したレイアウトが収まる最大の倍率。ただし古い Excel (lastEdited ≤ 6: 2016 以前の MSI 版) で
     *       保存したブックは、保存された倍率を使う。</li>
     * </ol>
     *
     * @param areas 印刷範囲 (この順にページを並べる)
     * @return ページの一覧
     */
    public List<Page> paginate(List<CellRangeAddress> areas) {
        if (ps.headings) {
            // 行列番号の印刷: 列の見出しの高さはシートの標準の行の高さ。行番号の列の幅は (桁数 + 0.75) × 標準フォントの数字の送り幅。
            // 桁数は最後の行の番号 (1 始まり) の桁数で、3 桁以下は同じ (Excel の印刷 exp80: 8 種類のフォント・大きさで 0.4pt 以内)
            int last = 0;
            for (CellRangeAddress a : areas) last = Math.max(last, a.getLastRow() + 1);
            int digits = Math.max(3, String.valueOf(last).length());
            hdrW = digitAdvance > 0 ? (digits + HDR_PAD_DIGITS) * digitAdvance : (HDR_W_MDW + HDR_DIGIT_MDW * (digits - 2)) * L.mdw;
            hdrH = L.defaultRowH;
        }
        double scale = ps.scale / 100.0;
        // 倍率を指定したシート: Excel 2016 以降は 100% のレイアウトのまま縮小する。
        // Excel 2010 で作られたブックは実体化に合う。PDF は最後に保存した Excel で作られることが多いので、
        // 古い Excel で保存したブックは実体化する
        boolean legacyApp = legacyExcel();
        if (!ps.fitToPage && !legacyApp && ps.scale != 100) L.linear = true;
        if (ps.fitToPage) {
            boolean stored = ps.scaleSet && ps.scale <= 100;
            int le = L.model().lastEdited;
            boolean oldBook = le > 0 && le <= 6;
            L.linear = FIT_REAL == 0;
            long lin = Math.round(fitScale(areas) * 100);
            int tol = oldBook ? 0 : FITLIN_TOL;
            L.linear = stored && ps.scale <= lin && ps.scale >= lin - tol;
            // 今の Excel は「ページに合わせる」のシートを、保存された倍率によらず計算し直した倍率で線形に縮小する。
            // --author-env は従来の判定
            if (!L.model().authorEnv && !legacyApp) {
                L.linear = true;
                ps.scale = (int) lin;
            }
            if (L.linear) {
                scale = ps.scale / 100.0;
            } else {
                double computed = fitScale(areas);
                scale = stored && oldBook ? ps.scale / 100.0 : computed;
            }
        }
        L.setScale(scale);
        List<Page> pages = new ArrayList<>();
        for (CellRangeAddress area : areas) pages.addAll(paginateArea(area, scale));
        return pages;
    }

    /**
     * 最後に保存した Excel が古い (縮小印刷でフォントを実体化し直す) か。既定は AppVersion ≤ 15 (2013 以前)。
     * x21 / webnew / webold: 判定なし 0.8169 / 0.8532 / 0.7775、AppVersion 0.7756 / 0.8533 / 0.9278、lastEdited ≤ 6 0.7655 / 0.8459 / 0.9285。
     */
    private boolean legacyExcel() {
        PrinterModel m = L.model();
        // 今の Excel は保存した版によらず線形に縮小する。版で分けるのは --author-env のときだけ
        if (!m.authorEnv) return false;
        return m.appVersion > 0 && m.appVersion <= 15;
    }

    /** 行列番号の印刷の見出しの大きさ (シートの pt)。 */
    private double hdrW, hdrH;
    /** 行番号の列の幅 (2 桁) と 1 桁あたりの幅の、標準フォントの数字の幅 (画素) に対する比 (Excel の印刷の実測)。 */
    static final double HDR_W_MDW = 2.808, HDR_DIGIT_MDW = 0.795;
    /** 行番号の列の幅の余白 (数字の送り幅の何文字分か)。 */
    static final double HDR_PAD_DIGITS = 0.75;
    /** 標準フォントの数字の送り幅 (pt)。0 なら最大数字幅 (mdw) から近似する。 */
    public double digitAdvance;

    /** 横に使える幅 (pt)。Excel は左右に少なくとも印刷可能領域の余白を残す。 */
    private double availW() {
        double mm = ps.minMargin;
        return Math.max(1, ps.paperW - Math.max(ps.left, mm) - Math.max(ps.right, mm) ) * fitShrink;
    }

    /**
     * 判定に使う大きさ (pt): 用紙をプリンタの画素に四捨五入し、余白を切り捨て、余裕 px を引く。既定の動きだけ。それ以外は avail を返す。
     *
     * @param avail 端数のままの使える大きさ (pt)
     * @param paper 用紙の幅か高さ (pt)
     * @param m0    片側の余白 (pt)
     * @param m1    反対側の余白 (pt)
     */
    private double intAvail(double avail, double paper, double m0, double m1, double px, int dpi) {
        if (px <= 0 || L.model().authorEnv) return avail;
        double k = dpi / 72.0;
        return (Math.floor(paper * k + 0.5) - Math.floor(m0 * k + 1e-6) - Math.floor(m1 * k + 1e-6) - px) / k;
    }

    /** 「ページに合わせる」の倍率を求めている間か。 */
    private boolean inFit = false;

    /** 「ページに合わせる」の倍率を求める間だけ、使える幅・高さに掛ける係数。 */
    private double fitShrink = 1;

    /** 縦に使える高さ (pt)。 */
    private double availH() {
        double mm = ps.minMargin;
        return Math.max(1, ps.paperH - Math.max(ps.top, mm) - Math.max(ps.bottom, mm)) * fitShrink;
    }

    /** 中央の式の内容に足す画素。 */
    static final int CENTER_ADD = 16;

    /** 中央に印刷するときの表の位置 (pt): 余白は切り捨て、内容に太線の分の画素を足した残りの半分を切り捨て、罫線の中心のずれを足す。 */
    private static double centerPx(double paper, double m0, double m1, double content, double scale) {
        long p = Math.round(Math.rint(paper / 0.12));
        long a = (long) Math.floor(m0 / 0.12 + 1e-6), b = (long) Math.floor(m1 / 0.12 + 1e-6);
        long c = Math.round(content / 0.12);
        long room = p - a - b - c - CENTER_ADD;
        return (a + Math.floorDiv(Math.max(room, -CENTER_ADD), 2L)) * 0.12 + posxOffset(scale);
    }

    /** 列の改ページで、整数の使える幅から引くプリンタの画素 (行の PAGE_PXH と同じ)。 */
    static final double PAGE_PXW = 8.0;

    
    /**
     * 「ページに合わせる」の倍率を求める。Excel は、その倍率で実体化したフォントによる自身のレイアウトが
     * 指定のページ数に収まる最大の倍率 (1% 刻み、10〜100%) を使う。左右の最小余白も効く ({@link #availW})。
     *
     * @return 倍率 (1.0 = 100%)
     */
    private double fitScale(List<CellRangeAddress> areas) {
        // 割合の余裕は --author-env だけ。既定の動きは画素の余裕 (FIT_PXW、FIT_PXH)
        double tight = !L.model().authorEnv ? 0 : FIT_TIGHT_DEFAULT;
        fitShrink = L.linear ? 1 - tight : 1;
                        inFit = true;
        try { return fitScale0(areas); } finally { inFit = false; fitShrink = 1; }
    }

    static final boolean DBG_FIT = System.getProperty("x2p.dbgfit") != null;
    /** 「ページに合わせる」の倍率を、実体化した寸法 (倍率ごとのフォントの標準行の比) で判定する (試験用。0 = 線形)。 */
    static final int FIT_REAL = Integer.getInteger("x2p.fitreal", 0);

    private double fitScale0(List<CellRangeAddress> areas) {
        int fw = ps.fitW, fh = ps.fitH;
        for (int s = 100; s >= 10; s--) {
            boolean ok = true;
            double sc = s / 100.0;
            L.setScale(sc);
            for (CellRangeAddress a : areas) {
                colOver = rowOver = false;
                int wide = colBands(a, sc).size();
                int tall = rowBands(a, sc).size();
                if (DBG_FIT && s >= 70 && s <= 95) System.err.println("FIT " + s + " area=" + a.formatAsString() + " wide=" + wide + " tall=" + tall + " colOver=" + colOver + " rowOver=" + rowOver + " availH=" + availH() + " y(last+1)=" + L.y(a.getLastRow() + 1) + " y(first)=" + L.y(a.getFirstRow()) + " hdrH=" + hdrH + " rh0-3=" + L.rh(0) + "," + L.rh(1) + "," + L.rh(2) + "," + L.rh(22) + " ratio=" + L.dbgInfo() + " y41=" + L.y(41) + " y40=" + L.y(40) + " y38=" + L.y(38));
                if ((fw > 0 && (wide > fw || colOver)) || (fh > 0 && (tall > fh || rowOver))) { ok = false; break; }
                // 丸める前の 100% の幅の合計から求めた倍率の四捨五入が上限 (Fableレポート B: 試験 fb-t2 の fit 78 件中 77 件)
                if (FIT_RAW_W && fw > 0) {
                    rawPass = true;
                    try {
                        if (colBands(a, sc).size() > fw) { ok = false; break; }
                    } finally {
                        rawPass = false;
                    }
                }
            }
            // 収まるかの判定は、実体化した寸法でも行う
            if (ok) return sc;
        }
        return 0.1;
    }

    /** 「ページに合わせる」の倍率を求めるときは、行タイトルの高さを途中から始まるページ (タイトル行より前から始まるページ) にも足す (夜間 n1801)。 */
    static final int TITLE_ALL = 1;
    /** 列タイトルも同じ: 倍率を求めるときは、どのページにも列タイトルの幅を足す (夜間 n2686)。 */
    static final boolean TITLE_COLS_FIT = true;

    /** 行タイトルの高さ (pt)。タイトル行より後から始まるページにだけ付く。 */
    private double titleRowsH(int bandStart) {
        CellRangeAddress t = ps.titleRows;
        if (t == null) return 0;
        if (bandStart <= t.getLastRow() && !(TITLE_ALL == 2 || (TITLE_ALL == 1 && inFit))) return 0;
        if (bandStart >= t.getFirstRow() && bandStart <= t.getLastRow() && !(FIT_TITLE_FIRST && inFit)) return 0;
        return (L.y(t.getLastRow() + 1) - L.y(t.getFirstRow()));
    }

    /** 「ページに合わせる」の倍率を求めるときは、タイトル行から始まる最初のページにも行タイトルの高さを足す (二重に数える。exp127 fit1x1_tt_tall・fit1x3_tt)。 */
    static final boolean FIT_TITLE_FIRST = true;

    /** 列タイトルの幅 (pt)。タイトル列より後から始まるページにだけ付く。 */
    private double titleColsW(int bandStart) {
        CellRangeAddress t = ps.titleCols;
        if (t == null) return 0;
        if (bandStart <= t.getLastColumn() && !(TITLE_COLS_FIT && inFit)) return 0;
        if (bandStart >= t.getFirstColumn() && bandStart <= t.getLastColumn() && !(TITLE_COLS_FIT && inFit)) return 0;
        return L.x(t.getLastColumn() + 1) - L.x(t.getFirstColumn());
    }

    /** 「ページに合わせる」では、ページ数が自動 (0) の方向の手動改ページだけが有効。 */
    private boolean manualRows() { return !ps.fitToPage || ps.fitH == 0; }

    private boolean manualCols() { return !ps.fitToPage || ps.fitW == 0; }

    /** 行をページの高さで区切る (手動改ページと行タイトルを考慮)。 */
    private List<Band> rowBands(CellRangeAddress a, double scale) {
        boolean manual = manualRows();
        List<Band> out = new ArrayList<>();
        double avail = availH() / scale * 1;
        if (fitShrink == 1) avail = intAvail(avail * scale, ps.paperH, ps.top, ps.bottom, inFit ? FIT_PXH : PAGE_PXH, ps.vdpi) / scale;
        avail -= hdrH;
        boolean exact = inFit && !L.model().authorEnv;   // 画素の余裕だけで判定する
        int start = a.getFirstRow();
        while (start <= a.getLastRow()) {
            double limit = avail - titleRowsH(start);
            double h = 0;
            int r = start;
            while (r <= a.getLastRow()) {
                if (r > start && manual && ps.rowBreaks.contains(r)) break;
                double rh = pageCeil(L.rh(r), scale, ps.vdpi);
                if (r > start && h + rh > limit + (exact ? 1e-6 : L.model().authorEnv && (scale < 1) ? -ROW_STRICT : 0.01)) {
                    break;
                }
                if (r == start && inFit && ONE_OVER > 1 && rh > limit + 1e-6) rowOver = true;
                h += rh;
                r++;
            }
            out.add(new Band(start, r - 1));
            start = r;
            if (maxPages > 0 && out.size() > maxPages) break;  // 上限を超えることは確定 (paginateArea で中止)
        }
        return out;
    }

    /** 線形に縮小するシートで、行高・列幅を縮小後のデバイス画素に切り上げる (シートの pt で返す)。 */
    private double pageCeil(double v, double scale, int dpi) {
        return v;
    }

    /** 倍率の上限の 2 回目の判定 (丸める前の幅と「倍率 − 0.5%」) の最中か。 */
    private boolean rawPass;
    static final boolean FIT_RAW_W = !"0".equals(System.getProperty("x2p.fitw"));

    /** 「ページに合わせる」の判定で、1 列 (1 行) だけで使える幅 (高さ) を超えたか。Excel はその列も収まるまで縮小する。 */
    private boolean colOver, rowOver;
    static final int ONE_OVER = 2;

    /** 列をページの幅で区切る (手動改ページと列タイトルを考慮)。 */
    private List<Band> colBands(CellRangeAddress a, double scale) {
        boolean manual = manualCols();
        List<Band> out = new ArrayList<>();
        double avail = availW() / scale;
        boolean raw = inFit && rawPass && L.linear && !L.model().authorEnv && fitShrink == 1;
        if (raw) avail = intAvail(availW(), ps.paperW, ps.left, ps.right, FIT_PXW, ps.hdpi) / (scale - 0.005);
        else if (inFit && fitShrink == 1) avail = intAvail(avail * scale, ps.paperW, ps.left, ps.right, FIT_PXW, ps.hdpi) / scale;
        else if (!inFit) avail = intAvail(avail * scale, ps.paperW, ps.left, ps.right, PAGE_PXW, ps.hdpi) / scale;
        avail -= hdrW;
        boolean exact = inFit && !L.model().authorEnv || !inFit && !L.model().authorEnv;   // 画素の余裕だけで判定する
        int start = a.getFirstColumn();
        while (start <= a.getLastColumn()) {
            double limit = avail - titleColsW(start);
            double w = 0;
            int c = start;
            while (c <= a.getLastColumn()) {
                if (c > start && manual && ps.colBreaks.contains(c)) break;
                double cw = raw ? L.cw100(c) : pageCeil(L.cw(c), scale, ps.hdpi);
                if (c > start && w + cw > limit + (exact ? 1e-6 : !L.model().authorEnv ? COL_TOL : 0.01)) break;
                if (c == start && inFit && ONE_OVER > 0 && cw > limit + 1e-6) colOver = true;
                w += cw;
                c++;
            }
            out.add(new Band(start, c - 1));
            start = c;
            if (maxPages > 0 && out.size() > maxPages) break;
        }
        return out;
    }

    /** 1 つの印刷範囲をページに分割する (ページの順序は設定の「左から右」「上から下」に従う)。 */
    private List<Page> paginateArea(CellRangeAddress a, double scale) {
        List<Band> rows = rowBands(a, scale);
        List<Band> cols = colBands(a, scale);
        long n = (long) rows.size() * cols.size();
        if (maxPages > 0 && made + n > maxPages) {
            throw new xlsx2pdf.LimitExceededException("ページ数が上限 (" + maxPages + ") を超えます (シート「" + L.sheet.getSheetName()
                    + "」の印刷範囲 " + a.formatAsString() + " が約 " + n + " ページ)。--max-pages で上限を上げるか、--no-limits で外せます", false);
        }
        made += n;
        List<Page> out = new ArrayList<>();
        if (ps.overThenDown) {
            for (Band rb : rows) for (Band cb : cols) addPage(out, a, rb, cb, scale);
        } else {
            for (Band cb : cols) for (Band rb : rows) addPage(out, a, rb, cb, scale);
        }
        return out;
    }

    static final boolean ZERO_MANUAL = true;

    /** 行・列の範囲からページを作って追加する (大きさが 0 のページ (すべて非表示) は作らない)。 */
    private void addPage(List<Page> out, CellRangeAddress a, Band rb, Band cb, double scale) {
        double bodyH = L.y(rb.b + 1) - L.y(rb.a);
        double bodyW = L.x(cb.b + 1) - L.x(cb.a);
        // 手動の改ページで始まる帯は、すべて非表示でも Excel は (空の) ページを印刷する (夜間 n2351: 改ページの右が非表示の列だけ)
        boolean manualStart = ZERO_MANUAL && (manualCols() && ps.colBreaks.contains(cb.a) && bodyH > 0
                || manualRows() && ps.rowBreaks.contains(rb.a) && bodyW > 0);
        if ((bodyH <= 0 || bodyW <= 0) && !manualStart) return;
        Page p = new Page();
        p.scale = scale;
        p.area = a;
        p.bodyR0 = rb.a; p.bodyR1 = rb.b; p.bodyC0 = cb.a; p.bodyC1 = cb.b;
        double th = titleRowsH(rb.a), tw = titleColsW(cb.a);
        CellRangeAddress tr = ps.titleRows, tc = ps.titleCols;
        if (th > 0 && tw > 0)
            p.segments.add(new Segment(tr.getFirstRow(), tr.getLastRow(), tc.getFirstColumn(), tc.getLastColumn(), 0, 0, false));
        if (th > 0) p.segments.add(new Segment(tr.getFirstRow(), tr.getLastRow(), cb.a, cb.b, tw, 0, false));
        if (tw > 0) p.segments.add(new Segment(rb.a, rb.b, tc.getFirstColumn(), tc.getLastColumn(), 0, th, false));
        p.segments.add(new Segment(rb.a, rb.b, cb.a, cb.b, tw, th, true));
        for (Segment sg : p.segments) sg.area = a;
        p.contentW = tw + bodyW + hdrW;
        p.contentH = th + bodyH + hdrH;
        p.hdrW = hdrW;
        p.hdrH = hdrH;
        double ox = Math.max(ps.left, ps.minMargin), oy = ps.top;
        boolean centerH = ps.centerH && availW() > p.contentW * scale;
        boolean posm = (L.linear || scale == 1);
        boolean posx = posm && !L.model().authorEnv;
        // Excel は余白を整数の画素に切り捨てる (全件の罫線の実測: 位置の差 = 余白の画素の端数)
        if (posx) {
            ox = Math.floor(ox / 0.12 + 1e-6) * 0.12;
            oy = Math.floor(oy / 0.12 + 1e-6) * 0.12;
        }
        if (centerH && posx) ox = centerPx(ps.paperW, Math.max(ps.left, ps.minMargin), Math.max(ps.right, ps.minMargin), p.contentW * scale, scale);
        else if (centerH) ox += (availW() - p.contentW * scale) / 2 + (posx ? -posxCenter(scale) : posm ? POSM_CENTER_DX_MODERN : 0);
        else if (posm && !L.model().authorEnv) ox += posxOffset(scale);
        else if (posm) ox += scale != 1 ? POSM_DX : POSM_DX100_MODERN;
        else if (scale != 1) ox += SCALED_DX;
        if (ps.centerV && posx) oy = centerPx(ps.paperH, ps.top, ps.bottom, p.contentH * scale, scale);
        else if (ps.centerV) oy += Math.max(0, (availH() - p.contentH * scale) / 2) - (posx ? posxCenter(scale) : 0);
        else if (posm && !L.model().authorEnv) oy += posxOffset(scale);
        else if (posm) oy += scale != 1 ? POSM_DX : POSM_DY100_MODERN;
        else if (scale != 1) oy += SCALED_DY;
        // 本文は見出しの分だけ右・下にずれる
        p.offsetX = Math.max(ox, ps.minMargin) + hdrW * scale;
        p.offsetY = Math.max(oy, ps.minMargin) + hdrH * scale;
        out.add(p);
    }
}
