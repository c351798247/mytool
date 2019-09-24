package cym.excel;


import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Created by Administrator on 2019/9/18.
 * Excel工具类 by JExcelAPI
 * 步骤：1.createExcel
 *      2.createSheet
 *      3.
 *      4.finish
 */
public class ExcelTool {

    public static final String EXCEL_DIR = System.getProperty("catalina.base").replace("file:\\", "") + File.separatorChar + "temp" + File.separatorChar + "excel" + File.separatorChar;
    public static void main(String[] args) throws IOException, WriteException {
        WritableWorkbook excel = ExcelTool.createExcel("2.xls");
        WritableSheet ws = ExcelTool.createSheet(excel, "first", 0);
//        for (int i = 0; i < 24; i++) {
//            ws.setRowView(i,500);
//            ws.setColumnView(i,3);
//        }
        setDataAndStyle(ws, "项目", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 1, 4, 1);
        setDataAndStyle(ws, "代码", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 2, 4, 2);
        setDataAndStyle(ws, "异地就医登记备案人数", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 3, 4, 3);
        setDataAndStyle(ws, "异地就医人数", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 4, 4, 4);
        setDataAndStyle(ws, "普通门（急）诊", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 5, 1, 9);
        setDataAndStyle(ws, "门诊大病（慢特病）", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 10, 1, 16);
        setDataAndStyle(ws, "住院", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 17, 1, 24);
        ExcelTool.finish(excel);
    }

    public static String getExcelFilePath(String csvFileSubName) {
//		LOG.error("catalina.base:" + System.getProperty("catalina.base"));
//		LOG.error("EXCEL_DIR:" + EXCEL_DIR);
//		LOG.error("sep:" + File.separatorChar);
        return new File(EXCEL_DIR, csvFileSubName).getPath();
    }
    /**
     * 创建Excel表格
     * @param fileName 文件名
     * @throws IOException
     * @throws WriteException
     */
    public static WritableWorkbook createExcel(String fileName) throws IOException {
        File file = new File(EXCEL_DIR, fileName);
        file.getParentFile().mkdirs();
        OutputStream os = new FileOutputStream(file.getPath());//取得输出流
        WritableWorkbook workbook = Workbook.createWorkbook(os);
//        setDataAndStyle(ws, "项目", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 1, 4, 1);
//        setDataAndStyle(ws, "代码", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 2, 4, 2);
//        setDataAndStyle(ws, "异地就医登记备案人数", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 3, 4, 3);
//        setDataAndStyle(ws, "异地就医人数", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 4, 4, 4);
//        setDataAndStyle(ws, "普通门（急）诊", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 5, 1, 9);
//        setDataAndStyle(ws, "门诊大病（慢特病）", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 10, 1, 16);
//        setDataAndStyle(ws, "住院", "宋体", 10, Border.NONE, Colour.GREY_25_PERCENT, true, 1, 17, 1, 24);
        return workbook;
    }

    /**
     *
     * 创建表单
     * @param workbook
     * @param sheetName
     * @param index 索引从0开始
     * @return
     */
    public static WritableSheet createSheet(WritableWorkbook workbook,String sheetName, int index) {
        return workbook.createSheet(sheetName, index);
    }
    /**
     * 完成导出
     * @param workbook
     * @throws IOException
     * @throws WriteException
     */
    public static void finish(WritableWorkbook workbook) throws IOException, WriteException {
        workbook.write();
        workbook.close();
    }
    /**
     *
     * @param sheet 表单
     * @param val   表格内值
     * @param FontName  字体名称
     * @param FontSize  字体大小
     * @param withoutBorder 取消的边框
     * @param backgroundColour 背景色
     * @param isMergeCells 是否合并单元格
     * @param startRow 起始行号 1开始
     * @param startCol 起始列号 1开始
     * @param endRow 终止行号 1开始
     * @param endCol 终止列号 1开始
     * @throws WriteException
     */
    public static void setDataAndStyle
    (WritableSheet sheet,String val,
     String FontName,int FontSize,
     Border withoutBorder,Colour backgroundColour,
     boolean isMergeCells,int startRow,int startCol,int endRow,int endCol)
            throws WriteException {
        WritableFont wfont = new WritableFont(WritableFont.createFont(FontName), FontSize);
        WritableCellFormat wc = new WritableCellFormat(wfont);

        wc.setAlignment(Alignment.CENTRE); // 设置水平居中
        wc.setVerticalAlignment(VerticalAlignment.CENTRE);//垂直居中
        wc.setWrap(true);//自动换行
        wc.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLUE); // 设置边框线
        switch (withoutBorder.getDescription()) {
            case "none" : wc.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.BLUE);break;//全部边框
            case "all" : wc.setBorder(Border.NONE, BorderLineStyle.THIN, Colour.BLUE);break;//无边框
            case "left" : wc.setBorder(Border.NONE, BorderLineStyle.THIN, Colour.BLUE);//去掉左边框
                wc.setBorder(Border.TOP, BorderLineStyle.THIN, Colour.BLUE);
                wc.setBorder(Border.RIGHT, BorderLineStyle.THIN, Colour.BLUE);
                wc.setBorder(Border.BOTTOM, BorderLineStyle.THIN, Colour.BLUE);break;
            case "right" : wc.setBorder(Border.NONE, BorderLineStyle.THIN, Colour.BLUE);//去掉右边框
                wc.setBorder(Border.TOP, BorderLineStyle.THIN, Colour.BLUE);
                wc.setBorder(Border.LEFT, BorderLineStyle.THIN, Colour.BLUE);
                wc.setBorder(Border.BOTTOM, BorderLineStyle.THIN, Colour.BLUE);break;
            default:break;
        }
        wc.setBackground(backgroundColour); // 设置单元格的背景颜色
        if (isMergeCells) {
            sheet.mergeCells(startCol-1, startRow-1, endCol-1, endRow-1);
        }
        Label label = new Label(startCol-1, startRow-1, val, wc);
        sheet.addCell(label);
    }
}
