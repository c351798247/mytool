package cym.decimal;

import java.math.BigDecimal;

/**
 * Created by Administrator on 2019/11/8.
 */
public class BigDecimalUtil {
    public static void main(String[] args) {
        BigDecimal b = new BigDecimal(1);
        System.out.println(add(null, "n"));
        System.out.println(add(null, "3"));
    }

    public static BigDecimal add(BigDecimal b1, Object o2) {
        try {
            if (b1 == null) {
                    b1 = new BigDecimal(0);
                if (o2 != null) {
                    if (o2 instanceof BigDecimal) {
                        b1 = (BigDecimal) o2;
                    } else {
                        b1 = new BigDecimal(o2.toString());
                    }
                }
            } else {
                if (o2 != null) {
                    if (o2 instanceof BigDecimal) {
                        b1 = b1.add(((BigDecimal) o2));
                    } else {
                        b1 = b1.add(new BigDecimal(o2.toString()));
                    }
                }
            }

        } catch (NumberFormatException e) {
            System.err.println("存在非数字字符");
            e.printStackTrace();
        }
        return b1;
    }
}
