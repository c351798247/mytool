package cym.string;

/**
 * @program:mytool
 * @author:CYM
 * @createTime:2019-11-03 00-03
 */
public class StringUtil {
    public static boolean isNotNullOrEmpty(String string) {
        if (string == null) {
            return false;
        }
        return !string.trim().isEmpty();

    }
}
