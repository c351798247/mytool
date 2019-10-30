package cym.date;

import com.sun.istack.internal.NotNull;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Administrator on 2019/10/30.
 */
public class DateTool {
    public static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public static String dateToString(@NotNull Date date) {
        return dateFormat.format(date);
    }
}
