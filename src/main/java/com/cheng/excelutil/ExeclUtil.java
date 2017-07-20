package com.cheng.excelutil;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 表格工具类
 *
 * @author chengys4
 *         2017-07-17 12:52
 **/
public class ExeclUtil {

    public static final String FILETYPE_XLSX = "xlsx";
    public static final String FILETYPE_XLS = "xls";
    public static final int MAX_PAGE_DATA = 150000;

    private ExeclUtil() {

    }

    /**
     * 读取 xls/xlsx 文件
     *
     * @param filePath xls/xlsx 路径
     * @return 读取的数据
     */
    public static List<Map<String, String>> read(String filePath) throws IOException {
        InputStream stream = new FileInputStream(filePath);
        return read(stream, getFileType(filePath));
    }

    /**
     * 读取 xls/xlsx 文件
     *
     * @param inputStream xls/xlsx 文件流
     * @return 读取的数据
     */
    private static List<Map<String, String>> read(InputStream inputStream, String fileType) {
        List<Map<String, String>> data = new ArrayList<>();
        Workbook wb = null;
        try {
            if (fileType.equals(FILETYPE_XLS)) {
                wb = new HSSFWorkbook(inputStream);
            } else if (fileType.equals(FILETYPE_XLSX)) {
                wb = new XSSFWorkbook(inputStream);
            } else {
                throw new FileTypeNotSupportExecption("您输入的excel格式不正确");
            }
            Sheet sheet1 = wb.getSheetAt(0);
            String[] strings = getColsName(sheet1);
            for (Row row : sheet1) {
                Map<String, String> item = new HashMap<>();
                for (int i = 0; i < strings.length; i++) {
                    Cell cell = row.getCell(i);
                    if (cell == null) {
                        item.put(strings[i], "");
                    } else {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        item.put(strings[i], cell.toString());
                    }
                }
                data.add(item);
            }
            data.remove(0);
            return data;
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                }
            }
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                }
            }
        }
    }

    /**
     * 得到数据表的头
     *
     * @param sheet
     * @return
     */
    public static String[] getColsName(Sheet sheet) {
        Row firstrow = sheet.getRow(sheet.getFirstRowNum());
        String[] colname = new String[firstrow.getLastCellNum()];
        for (int i = 0; i < colname.length; i++) {
            Cell cell = firstrow.getCell(i);
            if (cell == null) {
                colname[i] = "null";
            } else {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                colname[i] = cell.getStringCellValue();
            }
        }
        return colname;
    }

    /**
     * 获取文件类型
     *
     * @param filename
     * @return
     */
    public static String getFileType(String filename) {
        if (filename == null || filename.trim().length() == 0) {
            return null;
        }
        return filename.substring(filename.lastIndexOf(".") + 1, filename.length());
    }


    /**
     * 写入一个表格
     *
     * @param sheetVO
     * @return 表格
     */
    public static Workbook writeExecl(Workbook workbook, SheetVO sheetVO) {
        return writeSheet(workbook, sheetVO);
    }


    /**
     * 写入一个表格
     *
     * @param workbook
     * @param sheetVOs
     * @return
     */
    public static Workbook writeExecl(Workbook workbook, List<SheetVO> sheetVOs) {
        for (SheetVO sheetVO : sheetVOs) {
            writeSheet(workbook, sheetVO);
        }
        return workbook;
    }


    /**
     * 写入一个页签
     *
     * @param workbook
     * @param sheetVO
     * @return
     */
    private static Workbook writeSheet(Workbook workbook, SheetVO sheetVO) {
        if (sheetVO.getData().size() > MAX_PAGE_DATA) {
            throw new OutAllowRowNumException(" max size is " + MAX_PAGE_DATA);
        }
        return writeSheet(workbook, sheetVO.getSheetName(), sheetVO.getCellInfoVOs(), sheetVO.getData());
    }

    private static Workbook writeSheet(Workbook workbook, String sheetName, List<CellInfoVO> cellInfoVOs, List<Map<String, Object>> data) {

        Map<String, Map<String, String>> translaters = new HashMap<>();
        String[] excelTitles = new String[cellInfoVOs.size()];
        String[] excelCode = new String[cellInfoVOs.size()];
        for (int i = 0; i < cellInfoVOs.size(); i++) {
            excelTitles[i] = cellInfoVOs.get(i).getTitle();
            excelCode[i] = cellInfoVOs.get(i).getCode();
            translaters.put(cellInfoVOs.get(i).getCode(), cellInfoVOs.get(i).getTranslater());
        }
        return writeSheet(workbook, sheetName, excelCode, excelTitles, translaters, data);
    }

    /**
     * 写入一个页签
     *
     * @param workbook
     * @param sheetName
     * @param code
     * @param title
     * @param translaters
     * @param data
     * @return
     */
    private static Workbook writeSheet(Workbook workbook, String sheetName, String[] code, String[] title, Map<String, Map<String, String>> translaters, List<Map<String, Object>> data) {

        //创建页签
        Sheet sheet = workbook.createSheet(sheetName);
        //定义表格宽度
        int defWidth = 4000;
        int[] excelCellWidths = new int[code.length];
        Arrays.fill(excelCellWidths, defWidth);

        //表格样式
        CellStyle cellStyle = createDataStyle(workbook);

        //写表头
        // 创建标题
        writeTitleRow(sheet, title, excelCellWidths);

        //写数据
        int datasize = data.size();
        for (int rowIndex = 1; rowIndex <= datasize; rowIndex++) {
            //准备写入的数据
            Map<String, Object> item = data.get(rowIndex - 1);
            //翻译数据
            item = translateData(translaters, item);
            //写数据
            writeDataRow(sheet, cellStyle, item, rowIndex, code);
        }
        return workbook;

    }

    /**
     * 写一行表格
     */
    private static void writeDataRow(Sheet sheet, CellStyle cellStyle, Map<String, Object> item, int rowIndex, String[] code) {
        Row row = sheet.createRow(rowIndex);
        writeRowData(cellStyle, row, item, code);
    }


    /**
     * 创建表头
     *
     * @param sheet
     * @param titles
     * @param cellWidths
     */
    public static void writeTitleRow(Sheet sheet, String[] titles, int[] cellWidths) {
        // 创建行
        Row row = sheet.createRow(0);
        for (int i = 0; i < titles.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(titles[i]);
            sheet.setColumnWidth(i, cellWidths[i]);
        }
    }

    /**
     * 翻译数据
     *
     * @param translaters
     * @param item
     * @return
     */
    private static Map<String, Object> translateData(Map<String, Map<String, String>> translaters, Map<String, Object> item) {
        Set set1 = item.entrySet();
        Iterator it = set1.iterator();
        while (it.hasNext()) {
            Map.Entry entry = (Map.Entry) it.next();
            String key = (String) entry.getKey();
            Map<String, String> stringStringMap = translaters.get(key.toLowerCase());
            if (stringStringMap != null) {
                String translatervalue = stringStringMap.get(entry.getValue() == null ? "" : entry.getValue().toString());
                item.put(key, translatervalue);
            }
        }
        return item;
    }

    /**
     * 写入一行数据
     *
     * @param cellStyle
     * @param aRow
     * @param data
     * @param ticode
     */
    private static void writeRowData(CellStyle cellStyle, Row aRow, Map<String, Object> data, String[] ticode) {
        for (int i = 0; i < ticode.length; i++) {
            Cell cell = createCell(cellStyle, aRow, i);
            // 获取元素值
            Object value = data.get(ticode[i]);
            if (value instanceof Double) {
                cell.setCellValue((Double) value);
            } else if (value instanceof String) {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue((String) value);
            } else if (value instanceof Integer) {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Integer) value);
            } else if (value instanceof Timestamp) {
                cell.setCellValue(parseDate((Timestamp) value));
            } else if (value instanceof BigDecimal) {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(((BigDecimal) value).doubleValue());
            } else if (value instanceof Long) {
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue((Long) value);
            } else if (value instanceof Date) {
                cell.setCellValue(parseStr((Date) value));
            }
        }
    }

    private static String parseDate(Timestamp timestamp) {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        return format.format(timestamp);
    }

    private static String parseStr(Date datestr) {
        try {
            SimpleDateFormat e = new SimpleDateFormat("yyyy-MM-dd");
            return e.format(datestr);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 创建表格样式
     *
     * @param wb
     * @return
     */
    public static CellStyle createDataStyle(Workbook wb) {
        CellStyle numberCellStyle = wb.createCellStyle();// 创建单元格样式
        numberCellStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);// 指定单元格居中对齐
        numberCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 指定单元格垂直居中对齐
        numberCellStyle.setWrapText(true);// 指定当单元格内容显示不下时自动换行
        Font fontContent = wb.createFont();
        fontContent.setFontName("宋体");
        numberCellStyle.setFont(fontContent);
        numberCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        numberCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        numberCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        numberCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        return numberCellStyle;
    }

    /**
     * 创建单元格
     *
     * @param cellStyle
     * @param row
     * @param ci
     * @return
     */
    public static Cell createCell(CellStyle cellStyle, Row row, int ci) {
        Cell cell = row.createCell(ci);
        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
        cell.setCellStyle(cellStyle);
        return cell;
    }
}
