package com.github.biaosang;

public class HappyExcelException extends RuntimeException{

    public HappyExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public HappyExcelException(String message) {
        super(message);
    }

    public HappyExcelException(Throwable cause) {
        super(cause);
    }

}
