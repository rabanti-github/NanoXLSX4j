package ch.rabanti.nanoxlsx4j.utils;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

import java.time.Duration;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.temporal.ChronoUnit;
import java.util.Date;

public class DataUtils {
    private DataUtils() {
        // Do not instantiate
    }

    /**
     * First date that can be displayed by Excel. Real values before this date cannot be processed.
     */
    public static final  LocalDateTime FIRST_ALLOWED_EXCEL_DATE = LocalDateTime.of(1900, 1, 1, 0, 0, 0);

    /**
     * Last date that can be displayed by Excel. Real values after this date cannot be processed.
     */
    public static final  LocalDateTime LAST_ALLOWED_EXCEL_DATE = LocalDateTime.of(9999, 12, 31, 23, 59, 59);

    /**
     * All dates before this date are shifted in Excel by -1.0, since Excel assumes wrongly that the year 1900 is a leap year.<br />
     * See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     */
    public static final  LocalDateTime FIRST_VALID_EXCEL_DATE = LocalDateTime.of(1900, 3, 1, 0, 0, 0);

    private static final float COLUMN_WIDTH_ROUNDING_MODIFIER = 256f;
    private static final float SPLIT_WIDTH_MULTIPLIER = 12f;
    private static final float SPLIT_WIDTH_OFFSET = 0.5f;
    private static final float SPLIT_WIDTH_POINT_MULTIPLIER = 3f / 4f;
    private static final float SPLIT_POINT_DIVIDER = 20f;
    private static final float SPLIT_WIDTH_POINT_OFFSET = 390f;
    private static final float SPLIT_HEIGHT_POINT_OFFSET = 300f;
    private static final float ROW_HEIGHT_POINT_MULTIPLIER = 1f / 3f + 1f;

    private static final LocalDateTime ROOT_DATE = LocalDateTime.of(1899, 12, 30, 0, 0, 0);
    private static final long ROOT_MILLIS =
            ChronoUnit.MILLIS.between(
                    LocalDateTime.of(1, 1, 1, 0, 0),
                    LocalDateTime.of(1899, 12, 30, 0, 0)
            );


    /**
     * Method to convert a date or date and time into the internal Excel time format (OAdate)
     *
     * <p>Remarks: Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
     * from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
     * In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br />
     * See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br />
     * See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param date Date to process (time zone will be treated as UTC)
     * @return Date or date and time as number string
     * @throws FormatException Thrown if the passed date cannot be translated to the OADate format
     */
    public static String getOADateTimeString(Date date)
    {
        double d = getOADateTime(date);
        return ParserUtils.toString(d);
    }

    /**
     * Method to convert a date or date and time into the internal Excel time format (OAdate)
     *
     * <p>Remarks: Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
     * from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
     * In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br />
     * See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br />
     * See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param date Date to process (time zone will be treated as UTC)
     * @return Date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static double getOADateTime(Date date)
    {
        return getOADateTime(date, false);
    }

    /**
     * Method to convert a date or date and time into the internal Excel time format (OAdate)
     *
     * <p>Remarks: Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates
     * from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.
     * In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br />
     * See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br />
     * See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param skipCheck Flag to skip the validity check if set to true
     * @param date Date to process (time zone will be treated as UTC)
     * @return Date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static double getOADateTime(Date date, boolean skipCheck)
    {
        LocalDateTime dateValue = LocalDateTime.ofInstant(
                date.toInstant(),
                ZoneOffset.UTC
        );
        if (!skipCheck
                && (dateValue.isBefore(FIRST_ALLOWED_EXCEL_DATE)
                || dateValue.isAfter(LAST_ALLOWED_EXCEL_DATE))) {
            throw new FormatException(
                    "The date is not in a valid range for Excel. " +
                            "Dates before 1900-01-01 or after 9999-12-31 are not allowed."
            );
        }
        //Date dateValue = date;
        if (dateValue.isBefore(FIRST_VALID_EXCEL_DATE))
        {
            dateValue = dateValue.minusDays(1); // Fix of the leap-year-1900-error
        }
        double timeFraction = dateValue.toLocalTime().toSecondOfDay() / 86400d;

        long days = ChronoUnit.DAYS.between(
                ROOT_DATE.toLocalDate(),
                dateValue.toLocalDate()
        );
        return days + timeFraction;
    }

    /**
     * Method to convert a time into the internal Excel time format (OAdate without days)
     *
     * <p>Remarks: The time is represented by a OAdate without the date component but a possible number of total days</p>
     *
     * @param time Time to process. The date component of the timespan is converted to the total numbers of days
     * @return Time as number string
     */
    public static String getOATimeString(Duration time)
    {
        double d = getOATime(time);
        return ParserUtils.toString(d);
    }

    /// <summary>
    /// Method to convert a time into the internal Excel time format (OAdate without days)
    /// </summary>
    /// <param name="time">Time to process. The date component of the timespan is converted to the total numbers of days</param>
    /// <returns>Time as number</returns>
    /// \remark <remarks>The time is represented by a OAdate without the date component but a possible number of total days</remarks>
    public static double getOATime(Duration time)
    {
        return time.toSeconds() / 86400d;
    }
}
