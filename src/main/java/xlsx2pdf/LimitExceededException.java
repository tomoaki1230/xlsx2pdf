package xlsx2pdf;

/**
 * 変換の上限 (ページ数・画像の画素数・時間など) を超えたため変換を中止したことを表す。
 * 悪意のある小さなファイルでメモリや時間を使い切られないための上限で、オプションで緩めたり外したりできる
 * ({@link Converter.Options})。
 */
public class LimitExceededException extends RuntimeException {
    /** 時間の上限を超えたか (それ以外はページ数・画像などの大きさの上限)。 */
    private final boolean timeout;

    /**
     * @param message 何の上限を超えたか (緩めるオプションの名前を含める)
     * @param timeout 時間の上限か
     */
    public LimitExceededException(String message, boolean timeout) {
        super(message);
        this.timeout = timeout;
    }

    /** 時間の上限を超えたか。 */
    public boolean isTimeout() {
        return timeout;
    }
}
