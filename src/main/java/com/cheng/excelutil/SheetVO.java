package com.cheng.excelutil;

import java.util.List;
import java.util.Map;

/**
 * 页签数据
 */
public class SheetVO {
    private String sheetName ;
    private String[] title;
    private String[] code;
    private Map<String, Map<String, String>> translaters ;
    private List<Map<String,Object>> data ;
    public SheetVO(){

    }
    public SheetVO(String sheetName){
        this.sheetName = sheetName ;
    }
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String[] getTitle() {
        return title;
    }

    public void setTitle(String[] title) {
        this.title = title;
    }

    public String[] getCode() {
        return code;
    }

    public void setCode(String[] code) {
        this.code = code;
    }

    public Map<String, Map<String, String>> getTranslaters() {
        return translaters;
    }

    public void setTranslaters(Map<String, Map<String, String>> translaters) {
        this.translaters = translaters;
    }

    public List<Map<String, Object>> getData() {
        return data;
    }

    public void setData(List<Map<String, Object>> data) {
        this.data = data;
    }
}
