package xlsx2pdf;

/**
 * 必要な代替フォントが見つからないことを表す (同梱フォントの jar (xlsx2pdf-fonts.jar) がなく、PC にも代わりになるフォントがない)。
 */
public class FontUnavailableException extends RuntimeException {
    /** @param message 利用者向けの説明 (対処を含める) */
    public FontUnavailableException(String message) {
        super(message);
    }
}
