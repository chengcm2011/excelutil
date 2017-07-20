package com.cheng.excelutil;

import java.util.List;
import java.util.Map;

/**
 * 页签数据
 */
public class SheetVO {
    private String sheetName ="sheet01" ;
    private List<CellInfoVO> cellInfoVOs ;
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

    public List<CellInfoVO> getCellInfoVOs() {
        return cellInfoVOs;
    }

    public void setCellInfoVOs(List<CellInfoVO> cellInfoVOs) {
        this.cellInfoVOs = cellInfoVOs;
    }

    public List<Map<String, Object>> getData() {
        return data;
    }

    public void setData(List<Map<String, Object>> data) {
        this.data = data;
    }
}
