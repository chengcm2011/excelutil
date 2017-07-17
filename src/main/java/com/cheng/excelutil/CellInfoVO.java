package com.cheng.excelutil;

import java.util.Map;

/**
 *  导出的表格的表头信息
 */
public class CellInfoVO {
    private String title ;
    private String code ;
    private Map<String,String> translater ;
    private int width ;

    public CellInfoVO(String code, String title) {
        this.code =code ;
        this.title=title;
    }
    public CellInfoVO(String code, String title,Map<String,String> translater) {
        this.code =code ;
        this.title=title;
        this.translater = translater ;
    }
    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public int getWidth() {
        return width;
    }

    public Map<String, String> getTranslater() {
        return translater;
    }

    public void setTranslater(Map<String, String> translater) {
        this.translater = translater;
    }

    public void setWidth(int width) {
        this.width = width;
    }
}
