/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

import java.time.LocalTime;
import java.util.Calendar;
import java.util.Date;

/**
 * Class for shared used (static) methods
 *
 * @author Raphael Stoeckli
 */
public class Helper {

    // ### C O N S T A N T S ###
    private static final long ROOT_TICKS;

    private static final float COLUMN_WIDTH_ROUNDING_MODIFIER = 256f;
    private static final float SPLIT_WIDTH_MULTIPLIER = 12f;
    private static final float SPLIT_WIDTH_OFFSET = 0.5f;
    private static final float SPLIT_WIDTH_POINT_MULTIPLIER = 3f / 4f;
    private static final float SPLIT_POINT_DIVIDER = 20f;
    private static final float SPLIT_WIDTH_POINT_OFFSET = 390f;
    private static final float SPLIT_HEIGHT_POINT_OFFSET = 300f;
    private static final float ROW_HEIGHT_POINT_MULTIPLIER = 1f / 3f + 1f;

    /**
     * Minimum valid OAdate value (1900-01-01)
     */
    public static final double MIN_OADATE_VALUE = 0;
    /**
     * Maximum valid OAdate value (9999-12-31)
     */
    public static final double MAX_OADATE_VALUE = 2958465.9999d;

    /**
     * All dates before this date are shifted in Excel by -1.0, since Excel assumes wrongly that the year 1900 is a leap year.<br/>
     * See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     */
    public static final Date FIRST_VALID_EXCEL_DATE;

    static {
        Calendar rootCalendar = Calendar.getInstance();
        rootCalendar.set(1899, Calendar.DECEMBER, 30, 0, 0, 0);
        ROOT_TICKS = rootCalendar.getTimeInMillis();
        rootCalendar.set(1900, 3, 1);
        FIRST_VALID_EXCEL_DATE = rootCalendar.getTime();
    }

    private double d;

// ### S T A T I C   M E T H O D S ###    

    /**
     * Method to calculate the OA date (OLE automation) of the passed date.<br>
     * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException.<br/>
     * Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates from the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month.<br/>
     * In conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999 milliseconds).<br/>
     * See also: <a href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1"><br/>
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br/>
     * See also: <a href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year"><br/>
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     *
     * @param date Date to convert
     * @return OA date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static String getOADateString(Date date) {
        Date dateValue = date;
        Calendar dateCal = Calendar.getInstance();
        dateCal.setTime(date);
        if (date.before(FIRST_VALID_EXCEL_DATE)) {
            dateCal.add(Calendar.DAY_OF_WEEK, -1); // Fix of the leap-year-1900-error
        }
        long currentTicks = dateCal.getTimeInMillis();
        double d = ((double) (dateCal.get(Calendar.SECOND) + (dateCal.get(Calendar.MINUTE) * 60) + (dateCal.get(Calendar.HOUR_OF_DAY) * 3600)) / 86400) + Math.floor((currentTicks - ROOT_TICKS) / (86400000));
        if (d < MIN_OADATE_VALUE || d > MAX_OADATE_VALUE) {
            throw new FormatException("FormatException", "The date is not in a valid range for Excel. Dates before 1900-01-01 are not allowed.");
        }
        return Double.toString(d);
    }

    /**
     * Method to convert a time into the internal Excel time format (OAdate without days).<br>
     * The time is represented by a OAdate without the date component. A time range is between &gt;0.0 (00:00:00) and &lt;1.0 (23:59:59)
     *
     * @param time Time to process
     * @return Time as number
     * @throws FormatException Throws a FormatException if the passed time is invalid
     */
    public static String getOATimeString(LocalTime time) {
        try {
            int seconds = time.getSecond() + time.getMinute() * 60 + time.getHour() * 3600;
            double d = seconds / 86400d;
            return Double.toString(d);
        } catch (Exception e) {
            throw new FormatException("ConversionException", "The time could not be transformed into Excel format (OADate).", e);
        }
    }

    /**
     * Method to calculate a common Date from the OA date (OLE automation) format<br>
     * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException
     *
     * @param oaDate OA date number
     * @return Converted date
     */
    public static Date getDateFromOA(double oaDate) {
        double remainder = oaDate - (long) oaDate;
        double hours = remainder * 24;
        double minutes = (hours - (long) hours) * 60;
        double seconds = (minutes - (long) minutes) * 60;
        Calendar dateCal = Calendar.getInstance();
        dateCal.setTimeInMillis(ROOT_TICKS);
        dateCal.add(Calendar.DATE, (int) oaDate);
        dateCal.add(Calendar.HOUR, (int) hours);
        dateCal.add(Calendar.MINUTE, (int) minutes);
        dateCal.add(Calendar.SECOND, (int) seconds + 1); // Rounding error
        return dateCal.getTime();
    }

    /**
     * Calculates the internal width of a column in characters. This width is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
     *
     * @param columnWidth Target column width (displayed in Excel)
     * @return The internal column width in characters, used in worksheet XML documents
     * @apiNote The internal width deviates slightly from the column width, entered in Excel. Although internal, the default column width of 10 characters is visible in Excel as 10.71.
     * The deviation depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
     * In case of the width 10.0 and the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11),
     * the internal width is approximately 10.7142857 (rounded to 10.71).<br/> Note that the column height is not affected by this consideration.
     * The entered height in Excel is the actual height in the worksheet XML documents.<br/>
     * This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
     * See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a>
     */
    public static float getInternalColumnWidth(float columnWidth) {
        return getInternalColumnWidth(columnWidth, 7f, 5f);
    }

    /**
     * Calculates the internal width of a column in characters. This width is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
     *
     * @param columnWidth   Target column width (displayed in Excel)
     * @param maxDigitWidth Maximum digit with of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return The internal column width in characters, used in worksheet XML documents
     * @apiNote The internal width deviates slightly from the column width, entered in Excel. Although internal, the default column width of 10 characters is visible in Excel as 10.71.
     * The deviation depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
     * In case of the width 10.0 and the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11),
     * the internal width is approximately 10.7142857 (rounded to 10.71).<br/> Note that the column height is not affected by this consideration.
     * The entered height in Excel is the actual height in the worksheet XML documents.<br/>
     * This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
     * See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a>
     */
    public static float getInternalColumnWidth(float columnWidth, float maxDigitWidth, float textPadding) {
        if (columnWidth <= 0f || maxDigitWidth <= 0f) {
            return 0f;
        } else if (columnWidth <= 1f) {
            return (float) Math.floor((columnWidth * (maxDigitWidth + textPadding)) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) / COLUMN_WIDTH_ROUNDING_MODIFIER;
        } else {
            return (float) Math.floor((columnWidth * maxDigitWidth + textPadding) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) / COLUMN_WIDTH_ROUNDING_MODIFIER;
        }
    }

    /**
     * Calculates the internal height of a row. This height is used only in the XML documents of worksheets and is usually not exposed to the (Excel) end user
     *
     * @param rowHeight Target row height (displayed in Excel)
     * @return The internal row height which snaps to the nearest pixel
     * @apiNote The height is based on the calculated amount of pixels. One point are ~1.333 (1+1/3) pixels.
     * After the conversion, the number of pixels is rounded to the nearest integer and calculated back to points.<br/>
     * Therefore, the originally defined row height will slightly deviate, based on this pixel snap
     */
    public static float getInternalRowHeight(float rowHeight) {
        if (rowHeight <= 0f) {
            return 0f;
        }
        double heightInPixel = Math.round(rowHeight * ROW_HEIGHT_POINT_MULTIPLIER);
        return (float) heightInPixel / ROW_HEIGHT_POINT_MULTIPLIER;
    }

    /**
     * Calculates the internal width of a split pane in a worksheet. This width is used only in the XML documents of worksheets and is not exposed to the (Excel) end user
     *
     * @param width Target column(s) width (one or more columns, displayed in Excel)
     * @return The internal pane width, used in worksheet XML documents in case of worksheet splitting
     * @apiNote The internal split width is based on the width of one or more columns.
     * It also depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
     * See also {@link #getInternalColumnWidth(float, float, float)} for additional details.<br/>
     * This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
     * See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a><br/>
     */
    public static float getInternalPaneSplitWidth(float width) {
        return getInternalPaneSplitWidth(width, 7f, 5f);
    }

    /**
     * Calculates the internal width of a split pane in a worksheet. This width is used only in the XML documents of worksheets and is not exposed to the (Excel) end user
     *
     * @param width         Target column(s) width (one or more columns, displayed in Excel)
     * @param maxDigitWidth Maximum digit with of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return The internal pane width, used in worksheet XML documents in case of worksheet splitting
     * @apiNote The internal split width is based on the width of one or more columns.
     * It also depends on the maximum digit width of the default font, as well as its text padding and various constants.<br/>
     * See also {@link #getInternalColumnWidth(float, float, float)} for additional details.<br/>
     * This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br/>
     * See also: <a href="https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter 18.3.1.13</a><br/>
     * The two override parameters maxDigitWidth and textPadding probably don't have to be other than maxDigitWidth = 7f and textPadding = 5f.
     */
    public static float getInternalPaneSplitWidth(float width, float maxDigitWidth, float textPadding) {
        float pixels;
        if (width <= 1f) {
            pixels = (float) Math.floor(width / SPLIT_WIDTH_MULTIPLIER + SPLIT_WIDTH_OFFSET);
        } else {
            pixels = (float) Math.floor(width * maxDigitWidth + SPLIT_WIDTH_OFFSET) + textPadding;
        }
        float points = pixels * SPLIT_WIDTH_POINT_MULTIPLIER;
        return points * SPLIT_POINT_DIVIDER + SPLIT_WIDTH_POINT_OFFSET;
    }

    /**
     * Calculates the internal height of a split pane in a worksheet. This height is used only in the XML documents of worksheets and is not exposed to the (Excel) user
     *
     * @param height Target row(s) height (one or more rows, displayed in Excel)
     * @return The internal pane height, used in worksheet XML documents in case of worksheet splitting
     * @apiNote The internal split height is based on the height of one or more rows. It also depends on various constants.<br/>
     * This method is derived from the Perl implementation by John McNamara (<a href="https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)
     */
    public static float getInternalPaneSplitHeight(float height) {
        return (float) Math.floor(SPLIT_POINT_DIVIDER * height + SPLIT_HEIGHT_POINT_OFFSET);
    }

    /**
     * Method of a string to check whether its reference is null or the content is empty
     *
     * @param value value / reference to check
     * @return True if the passed parameter is null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String value) {
        if (value == null) {
            return true;
        }
        return value.isEmpty();
    }


}
