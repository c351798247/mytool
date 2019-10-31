package cym.date;

import com.sun.istack.internal.NotNull;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Administrator on 2019/10/30.
 */
public class DateTool {
    private static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public static String dateToString(@NotNull Date date) {
        try {
            return dateFormat.format(date);
        } catch (Exception e) {
            return "日期转换异常";
        }
    }
    public static String dateToString(@NotNull Date date, String format) {
        String dateString = "";
        if (format == null) {
            dateString = dateFormat.format(date);
        } else {
            try {
                dateString = new SimpleDateFormat(format).format(date);
            } catch (Exception e) {
                dateString = "日期转换异常";
            }
        }
        return dateString;
    }

    public static void main(String[] args) {
        System.out.println(DateTool.dateToString(null));
    }
}
