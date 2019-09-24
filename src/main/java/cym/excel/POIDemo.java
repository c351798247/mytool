package cym.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2019/9/19.
 */
public class POIDemo {
    public static void main(String[] args) throws IOException {


    }

    /**
     *
     * @param sheet     sheet
     * @param cellStyle 单元格样式
     * @param value     值
     * @param isMarge   是否合并单元格
     * @param startRow  开始行
     * @param endRow    结束行
     * @param startCol  开始列
     * @param endCol    结束列
     */
    public static void setValue(Sheet sheet,CellStyle cellStyle,String value, boolean isMarge, int startRow, int endRow, int startCol, int endCol) {
        //判断是否合并单元格
        if (isMarge) {
            for (int i = startRow; i <= endRow; i++) {
                //获取行
                Row row = sheet.getRow(i) == null ? sheet.createRow(i) : sheet.getRow(i);
                for (int j = startCol; j <= endCol; j++) {
                    //获取单元格
                    Cell cell = row.createCell(j);
                    //设置样式
                    cell.setCellStyle(cellStyle);
                }
            }
            //合并单元格
            sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCol, endCol));
        } else {
            //获取行
            Row row = sheet.getRow(startRow) == null ? sheet.createRow(startRow) : sheet.getRow(startRow);
            //获取单元格
            Cell cell = row.createCell(startCol);
            //设置样式
            cell.setCellStyle(cellStyle);
        }
        //设置值
        sheet.getRow(startRow).getCell(startCol).setCellValue(value);
    }

    /**
     *
     * @param workbook  workbook
     * @param fontName  字体名称
     * @param fontSize  字体大小
     * @return
     */
    public static Font createFont(SXSSFWorkbook workbook,String fontName, short fontSize) {
        Font font = workbook.createFont();
        font.setFontName(fontName);//字体名称
        font.setFontHeightInPoints(fontSize);//字体大小
        return font;
    }

    /**
     *
     * @param workbook
     * @param font  字体
     * @param borderState   边框状态:0无边框
     *                              1全边框
     *                              2无左边框
     *                              3无右边框
     *                              4无上边框
     *                              5无下边框
     * @param backgroundColor 背景色
     * @return
     */
    public static CellStyle createCellStyle(SXSSFWorkbook workbook,Font font,int borderState,short backgroundColor) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);        //设置字体
        style.setWrapText(true);    //自动换行
        style.setAlignment(CellStyle.ALIGN_CENTER);         //水平居中
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); //垂直居中

        switch (borderState) {
            case 1 : style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                style.setBorderRight(HSSFCellStyle.BORDER_THIN);break;
            case 0 :style.setBorderBottom(HSSFCellStyle.BORDER_NONE);
                style.setBorderLeft(HSSFCellStyle.BORDER_NONE);
                style.setBorderTop(HSSFCellStyle.BORDER_NONE);
                style.setBorderRight(HSSFCellStyle.BORDER_NONE);break;
            case 2 : style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                style.setBorderLeft(HSSFCellStyle.BORDER_NONE);
                style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                style.setBorderRight(HSSFCellStyle.BORDER_THIN);break;
            case 3 : style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                style.setBorderRight(HSSFCellStyle.BORDER_NONE);break;
            case 4 :style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                style.setBorderTop(HSSFCellStyle.BORDER_NONE);
                style.setBorderRight(HSSFCellStyle.BORDER_THIN);break;
            case 5 : style.setBorderBottom(HSSFCellStyle.BORDER_NONE);
                style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                style.setBorderRight(HSSFCellStyle.BORDER_THIN);break;
        }
        //边框颜色
        style.setLeftBorderColor(HSSFColor.ROYAL_BLUE.index);
        style.setRightBorderColor(HSSFColor.ROYAL_BLUE.index);
        style.setBottomBorderColor(HSSFColor.ROYAL_BLUE.index);
        style.setTopBorderColor(HSSFColor.ROYAL_BLUE.index);
        style.setFillForegroundColor(backgroundColor);// 设置单元格的背景颜色．
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return style;
    }


    public static void finished(SXSSFWorkbook workbook,String fileName) {
        try{
            FileOutputStream fos = new FileOutputStream(fileName);
            workbook.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
