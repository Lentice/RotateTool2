package lahome.rotateTool.Util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Date;

public class DateUtil {

    private static final String LOCAL_DATE_PATTERN = "yyyy/MM/dd";

    private static final DateTimeFormatter LOCAL_DATE_FORMATTER =
            DateTimeFormatter.ofPattern(LOCAL_DATE_PATTERN);

    private static SimpleDateFormat DATE_FORMATTER = new SimpleDateFormat("yyyy/MM/dd");

    public static String format(Date date) {

        return DATE_FORMATTER.format(date);
    }

    public static Date parse(String dateString) {
        try {
        	return DATE_FORMATTER.parse(dateString);
        } catch (ParseException e) {
            return null;
        }
    }

    public static boolean validDate(String dateString) {
        // Try to parse the String.
        return DateUtil.parse(dateString) != null;
    }

    public static String formatFromLocalDate(LocalDate date) {
        if (date == null) {
            return null;
        }
        return LOCAL_DATE_FORMATTER.format(date);
    }

    public static LocalDate parseToLocalDate(String dateString) {
        try {
            return LOCAL_DATE_FORMATTER.parse(dateString, LocalDate::from);
        } catch (DateTimeParseException e) {
            return null;
        }
    }

    public static boolean validLocalDate(String dateString) {
        return DateUtil.parseToLocalDate(dateString) != null;
    }

    public static Date asDate(LocalDate localDate) {
        return java.sql.Date.valueOf(localDate);
    }

    public static Date asDate(LocalDateTime localDateTime) {
        return Date.from(localDateTime.atZone(ZoneId.systemDefault()).toInstant());
    }

    public static LocalDate asLocalDate(Date date) {
        return Instant.ofEpochMilli(date.getTime()).atZone(ZoneId.systemDefault()).toLocalDate();
    }

    public static LocalDateTime asLocalDateTime(Date date) {
        return Instant.ofEpochMilli(date.getTime()).atZone(ZoneId.systemDefault()).toLocalDateTime();
    }
}
