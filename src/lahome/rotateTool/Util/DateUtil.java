package lahome.rotateTool.Util;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtil {

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
}
