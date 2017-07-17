package com.cheng.excelutil;

/**
 * 超过最大行异常
 *
 * @author chengys4
 *         2017-07-17 18:02
 **/
public class OutAllowRowNumException extends RuntimeException {
    public OutAllowRowNumException(String message) {
        super(message);
    }
}
