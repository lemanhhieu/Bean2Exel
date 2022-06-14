package bean2Excel;

public class Bean2ExcelException extends RuntimeException {
    public Bean2ExcelException() {
    }

    public Bean2ExcelException(String message) {
        super(message);
    }

    public Bean2ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public Bean2ExcelException(Throwable cause) {
        super(cause);
    }

    public Bean2ExcelException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
