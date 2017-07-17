package com.cheng.excelutil;

/**
 * 格式不支持
 *
 * @author chengys4
 *         2017-07-17 18:08
 **/
public class FileTypeNotSupportExecption extends RuntimeException {
    public FileTypeNotSupportExecption(String message) {
        super(message);
    }
}
