package api.xl.util;

import java.time.LocalDate;
import java.time.temporal.ChronoUnit;

public final class XLDateUtil {
    private static final LocalDate startDate = LocalDate.of(1900, 1, 1);
    public static LocalDate toDate(String value) {
        double dias = Double.parseDouble(value);
        return startDate.plusDays((long) dias - 1);
    }

    public static String toString(LocalDate date) {

        long dias = ChronoUnit.DAYS.between(startDate, date)+1;
        return Long.toString(dias);
    }

    public static boolean isDate(String value) {
        if (value == null || !value.matches("[0-9]+(\\.[0-9]+)?")) {
            return false;
        }
        double days = Double.parseDouble(value);
        if (days <= 0) {
            return false;
        }
        LocalDate currentDate = LocalDate.now();
        return !toDate(value).isAfter(currentDate);
    }
}
