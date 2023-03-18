/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.time.Duration;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalAccessor;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

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
	 * Minimum valid OAdate value (1900-01-01). However, Excel displays this value
	 * as 1900-01-00 (day zero)
	 */
	public static final double MIN_OADATE_VALUE = 0;
	/**
	 * Maximum valid OAdate value (9999-12-31)
	 */
	public static final double MAX_OADATE_VALUE = 2958465.999988426d;
	/**
	 * First date that can be displayed by Excel. Values before this date cannot be
	 * processed.
	 */
	public static final Date FIRST_ALLOWED_EXCEL_DATE;

	/**
	 * Last date that can be displayed by Excel. Values after this date cannot be
	 * processed.
	 */
	public static final Date LAST_ALLOWED_EXCEL_DATE;

	/**
	 * All dates before this date are shifted in Excel by -1.0, since Excel assumes
	 * wrongly that the year 1900 is a leap year.<br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
	 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
	 */
	public static final Date FIRST_VALID_EXCEL_DATE;

	static {
		Calendar rootCalendar = Calendar.getInstance();
		rootCalendar.set(1899, Calendar.DECEMBER, 30, 0, 0, 0);
		rootCalendar.set(Calendar.MILLISECOND, 0);
		ROOT_TICKS = rootCalendar.getTimeInMillis();
		rootCalendar.set(1900, Calendar.MARCH, 1);
		FIRST_VALID_EXCEL_DATE = rootCalendar.getTime();
		rootCalendar.set(1900, Calendar.JANUARY, 1, 0, 0, 0);
		FIRST_ALLOWED_EXCEL_DATE = rootCalendar.getTime();
		rootCalendar.set(9999, Calendar.DECEMBER, 31, 23, 59, 59);
		LAST_ALLOWED_EXCEL_DATE = rootCalendar.getTime();
	}

	// ### C O N S T R U C T O R S ###
	private Helper() {
		// Prevents class instantiation
	}

	// ### S T A T I C M E T H O D S ###

	/**
	 * Method to calculate the OA date (OLE automation) of the passed date.<br>
	 * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond
	 * this date cannot be handled by Excel under normal circumstances and will
	 * throw a FormatException.<br>
	 * Excel assumes wrongly that the year 1900 is a leap year. There is a gap of
	 * 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates from
	 * the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the
	 * minimum valid date as 1900-01-00, although 0 is not a valid description for a
	 * day of month.<br>
	 * In conformance to the OAdate specifications, the maximum valid date is
	 * 9999-12-31 23:59:59 (plus 999 milliseconds).<br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1"><br>
	 * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year"><br>
	 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
	 *
	 * @param date
	 *            Date to convert
	 * @return OA date or date and time as number string
	 * @throws FormatException
	 *             Throws a FormatException if the passed date cannot be translated
	 *             to the OADate format
	 */
	public static String getOADateString(Date date) {
		double d = getOADate(date);
		return Double.toString(d);
	}

	/**
	 * Method to calculate the OA date (OLE automation) of the passed date.<br>
	 * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond
	 * this date cannot be handled by Excel under normal circumstances and will
	 * throw a FormatException.<br>
	 * Excel assumes wrongly that the year 1900 is a leap year. There is a gap of
	 * 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates from
	 * the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the
	 * minimum valid date as 1900-01-00, although 0 is not a valid description for a
	 * day of month.<br>
	 * In conformance to the OAdate specifications, the maximum valid date is
	 * 9999-12-31 23:59:59 (plus 999 milliseconds).<br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1"><br>
	 * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year"><br>
	 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
	 *
	 * @param date
	 *            Date to convert
	 * @return OA date or date and time as number
	 * @throws FormatException
	 *             Throws a FormatException if the passed date cannot be translated
	 *             to the OADate format
	 */
	public static double getOADate(Date date) {
		return getOADate(date, false);
	}

	/**
	 * Method to calculate the OA date (OLE automation) of the passed date.<br>
	 * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond
	 * this date cannot be handled by Excel under normal circumstances and will
	 * throw a FormatException.<br>
	 * Excel assumes wrongly that the year 1900 is a leap year. There is a gap of
	 * 1.0 between 1900-02-28 and 1900-03-01. This method corrects all dates from
	 * the first valid date (1900-01-01) to 1900-03-01. However, Excel displays the
	 * minimum valid date as 1900-01-00, although 0 is not a valid description for a
	 * day of month.<br>
	 * In conformance to the OAdate specifications, the maximum valid date is
	 * 9999-12-31 23:59:59 (plus 999 milliseconds).<br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1"><br>
	 * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br>
	 * See also: <a href=
	 * "https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year"><br>
	 * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
	 *
	 * @param date
	 *            Date to convert
	 * @param skipCheck
	 *            If true, the validity check is skipped</param>
	 * @return OA date or date and time as number
	 * @throws FormatException
	 *             Throws a FormatException if the passed date cannot be translated
	 *             to the OADate format
	 */
	public static double getOADate(Date date, boolean skipCheck) {
		if (date == null) {
			throw new FormatException("The date cannot be null");
		}
		if (!skipCheck && (date.before(FIRST_ALLOWED_EXCEL_DATE) || date.after(LAST_ALLOWED_EXCEL_DATE))) {
			throw new FormatException("The date is not in a valid range for Excel. Dates before 1900-01-01 are not allowed.");
		}
		Calendar dateCal = Calendar.getInstance();
		dateCal.setTime(date);
		if (date.before(FIRST_VALID_EXCEL_DATE)) {
			dateCal.add(Calendar.DAY_OF_WEEK, -1); // Fix of the leap-year-1900-error
		}
		long currentTicks = dateCal.getTimeInMillis();
		return ((dateCal.get(Calendar.SECOND) + (dateCal.get(Calendar.MINUTE) * 60) + (dateCal.get(Calendar.HOUR_OF_DAY) * 3600)) / 86400d) +
				Math.floor((currentTicks - ROOT_TICKS) / (86400000d));
	}

	/**
	 * Method to convert a duration into the internal Excel time format (OAdate
	 * without days).<br>
	 * The time is represented by a OAdate without the date component but a possible
	 * number of total days
	 *
	 * @param time
	 *            Time to process
	 * @return Time as number string
	 * @throws FormatException
	 *             Throws a FormatException if the passed time is invalid
	 */
	public static String getOATimeString(Duration time) {
		double d = getOATime(time);
		return Double.toString(d);
	}

	/**
	 * Method to convert a duration into the internal Excel time format (OAdate
	 * without days).<br>
	 * The time is represented by a OAdate without the date component but a possible
	 * number of total days
	 *
	 * @param time
	 *            Time to process
	 * @return Time as number
	 * @throws FormatException
	 *             Throws a FormatException if the passed time is invalid
	 */
	public static double getOATime(Duration time) {
		try {
			long seconds = time.get(ChronoUnit.SECONDS);
			return (double) seconds / 86400d;
		}
		catch (Exception e) {
			throw new FormatException("The time could not be transformed into Excel format (OADate).", e);
		}
	}

	/**
	 * Method to calculate a common Date from the OA date (OLE automation)
	 * format<br>
	 * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond
	 * this date cannot be handled by Excel under normal circumstances and will
	 * throw a FormatException
	 *
	 * @param oaDate
	 *            OA date number
	 * @return Converted date
	 * @implNote Numbers that represents dates before 1900-03-01 (number of days
	 *           since 1900-01-01 = 60) are automatically modified. Until 1900-03-01
	 *           is 1.0 added to the number to get the same date, as displayed in
	 *           Excel. The reason for this is a bug in Excel. See also: <a href=
	 *           "https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year"><br>
	 *           https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
	 */
	public static Date getDateFromOA(double oaDate) {
		if (oaDate < 60) {
			oaDate = oaDate + 1;
		}
		double remainder = oaDate - (long) oaDate;
		double hours = remainder * 24;
		double minutes = (hours - (long) hours) * 60;
		double seconds = (minutes - (long) minutes) * 60;
		seconds = Math.round(seconds);
		Calendar dateCal = Calendar.getInstance();
		dateCal.setTimeInMillis(ROOT_TICKS);
		dateCal.add(Calendar.DATE, (int) oaDate);
		dateCal.add(Calendar.HOUR, (int) hours);
		dateCal.add(Calendar.MINUTE, (int) minutes);
		dateCal.add(Calendar.SECOND, (int) seconds);
		return dateCal.getTime();
	}

	/**
	 * Method to parse a {@link Duration} object from a string and a pattern that is
	 * allied to a {@link DateTimeFormatter} instance
	 *
	 * @param timeString
	 *            String to parse
	 * @param pattern
	 *            Pattern as string
	 * @param locale
	 *            Locale of the formatter that is created from the pattern
	 * @return Duration object
	 * @throws FormatException
	 *             thrown if the time could not be parsed or of the pattern was
	 *             invalid
	 * @apiNote Supported formatting tokens are all time-related patterns like 'HH',
	 *          'mm', 'ss'. To represent the number of days, the pattern 'n' is
	 *          used. This deviates from the actual definition of 'n' which would be
	 *          nanoseconds of the second. date-related patterns like 'dd', 'MM' or
	 *          'yyyy' are not supported yet for time parsing
	 */
	public static Duration parseTime(String timeString, String pattern, Locale locale) {
		if (isNullOrEmpty(timeString) || isNullOrEmpty(pattern)) {
			throw new FormatException("The pattern is not valid");
		}
		try {
			DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern).withLocale(locale);
			return parseTime(timeString, formatter);
		}
		catch (Exception ex) {
			throw new FormatException("The time could not be parsed: " + ex.getMessage(), ex);
		}
	}

	/**
	 * Method to parse a {@link Duration} object from a string and a
	 * {@link DateTimeFormatter} instance
	 *
	 * @param timeString
	 *            String to parse
	 * @param formatter
	 *            Formatter to apply the parsing
	 * @return Duration object
	 * @throws FormatException
	 *             thrown if the time could not be parsed or of the formatter was
	 *             invalid
	 * @apiNote Note that the formatter may consider locales. Parsing may fail e.g.
	 *          with "AM/PM" when not on {@link Locale#US}. Supported formatting
	 *          tokens are all time-related patterns like 'HH', 'mm', 'ss'. To
	 *          represent the number of days, the pattern 'n' is used. This deviates
	 *          from the actual definition of 'n' which would be nanoseconds of the
	 *          second. date-related patterns like 'dd', 'MM' or 'yyyy' are not
	 *          supported yet for time parsing
	 */
	public static Duration parseTime(String timeString, DateTimeFormatter formatter) {
		if (isNullOrEmpty(timeString) || formatter == null) {
			throw new FormatException("Either no time value or formatter was defined and no time can be parsed");
		}
		try {
			TemporalAccessor accessor = formatter.parse(timeString);
			long days = 0;
			long hours = 0;
			long minutes = 0;
			long seconds = 0;
			if (accessor.isSupported(ChronoField.HOUR_OF_DAY)) {
				hours = accessor.get(ChronoField.HOUR_OF_DAY);
			}
			if (accessor.isSupported(ChronoField.MINUTE_OF_HOUR)) {
				minutes = accessor.get(ChronoField.MINUTE_OF_HOUR);
			}
			if (accessor.isSupported(ChronoField.SECOND_OF_MINUTE)) {
				seconds = accessor.get(ChronoField.SECOND_OF_MINUTE);
			}
			if (accessor.isSupported(ChronoField.NANO_OF_SECOND)) {
				days = accessor.get(ChronoField.NANO_OF_SECOND);
			}
			Duration duration = Duration.ZERO;
			duration = duration.plusDays(days);
			duration = duration.plusHours(hours);
			duration = duration.plusMinutes(minutes);
			duration = duration.plusSeconds(seconds);
			return duration;
		}
		catch (Exception ex) {
			throw new FormatException("The time could not be parsed: " + ex.getMessage(), ex);
		}
	}

	/**
	 * Calculates the internal width of a column in characters. This width is used
	 * only in the XML documents of worksheets and is usually not exposed to the
	 * (Excel) end user
	 *
	 * @param columnWidth
	 *            Target column width (displayed in Excel)
	 * @return The internal column width in characters, used in worksheet XML
	 *         documents
	 * @apiNote The internal width deviates slightly from the column width, entered
	 *          in Excel. Although internal, the default column width of 10
	 *          characters is visible in Excel as 10.71. The deviation depends on
	 *          the maximum digit width of the default font, as well as its text
	 *          padding and various constants.<br>
	 *          In case of the width 10.0 and the default digit width 7.0, as well
	 *          as the padding 5.0 of the default font Calibri (size 11), the
	 *          internal width is approximately 10.7142857 (rounded to 10.71).<br>
	 *          Note that the column height is not affected by this consideration.
	 *          The entered height in Excel is the actual height in the worksheet
	 *          XML documents.<br>
	 *          This method is derived from the Perl implementation by John McNamara
	 *          (<a href=
	 *          "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br>
	 *          See also: <a href=
	 *          "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376,
	 *          Part 1, Chapter 18.3.1.13</a>
	 */
	public static float getInternalColumnWidth(float columnWidth) {
		return getInternalColumnWidth(columnWidth, 7f, 5f);
	}

	/**
	 * Calculates the internal width of a column in characters. This width is used
	 * only in the XML documents of worksheets and is usually not exposed to the
	 * (Excel) end user
	 *
	 * @param columnWidth
	 *            Target column width (displayed in Excel)
	 * @param maxDigitWidth
	 *            Maximum digit with of the default font (default is 7.0 for
	 *            Calibri, size 11)
	 * @param textPadding
	 *            Text padding of the default font (default is 5.0 for Calibri, size
	 *            11)
	 * @return The internal column width in characters, used in worksheet XML
	 *         documents
	 * @throws FormatException
	 *             thrown if the column width is out of range
	 * @apiNote The internal width deviates slightly from the column width, entered
	 *          in Excel. Although internal, the default column width of 10
	 *          characters is visible in Excel as 10.71. The deviation depends on
	 *          the maximum digit width of the default font, as well as its text
	 *          padding and various constants.<br>
	 *          In case of the width 10.0 and the default digit width 7.0, as well
	 *          as the padding 5.0 of the default font Calibri (size 11), the
	 *          internal width is approximately 10.7142857 (rounded to 10.71).<br>
	 *          Note that the column height is not affected by this consideration.
	 *          The entered height in Excel is the actual height in the worksheet
	 *          XML documents.<br>
	 *          This method is derived from the Perl implementation by John McNamara
	 *          (<a href=
	 *          "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br>
	 *          See also: <a href=
	 *          "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376,
	 *          Part 1, Chapter 18.3.1.13</a>
	 */
	public static float getInternalColumnWidth(float columnWidth, float maxDigitWidth, float textPadding) {
		if (columnWidth < Worksheet.MIN_COLUMN_WIDTH || columnWidth > Worksheet.MAX_COLUMN_WIDTH) {
			throw new FormatException("The column width " +
					columnWidth +
					" is not valid. The valid range is between " +
					Worksheet.MIN_COLUMN_WIDTH +
					" and " +
					Worksheet.MAX_COLUMN_WIDTH);
		}
		if (columnWidth <= 0f || maxDigitWidth <= 0f) {
			return 0f;
		}
		else if (columnWidth <= 1f) {
			return (float) Math.floor((columnWidth * (maxDigitWidth + textPadding)) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) /
					COLUMN_WIDTH_ROUNDING_MODIFIER;
		}
		else {
			return (float) Math.floor((columnWidth * maxDigitWidth + textPadding) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) /
					COLUMN_WIDTH_ROUNDING_MODIFIER;
		}
	}

	/**
	 * Calculates the internal height of a row. This height is used only in the XML
	 * documents of worksheets and is usually not exposed to the (Excel) end user
	 *
	 * @param rowHeight
	 *            Target row height (displayed in Excel)
	 * @return The internal row height which snaps to the nearest pixel
	 * @throws FormatException
	 *             thrown if the row height is out of range
	 * @apiNote The height is based on the calculated amount of pixels. One point
	 *          are ~1.333 (1+1/3) pixels. After the conversion, the number of
	 *          pixels is rounded to the nearest integer and calculated back to
	 *          points.<br>
	 *          Therefore, the originally defined row height will slightly deviate,
	 *          based on this pixel snap
	 */
	public static float getInternalRowHeight(float rowHeight) {
		if (rowHeight < Worksheet.MIN_ROW_HEIGHT || rowHeight > Worksheet.MAX_ROW_HEIGHT) {
			throw new FormatException("The row height " +
					rowHeight +
					" is not valid. The valid range is between " +
					Worksheet.MIN_ROW_HEIGHT +
					" and " +
					Worksheet.MAX_ROW_HEIGHT);
		}
		if (rowHeight == 0f) {
			return 0f;
		}
		double heightInPixel = Math.round(rowHeight * ROW_HEIGHT_POINT_MULTIPLIER);
		return (float) heightInPixel / ROW_HEIGHT_POINT_MULTIPLIER;
	}

	/**
	 * Calculates the internal width of a split pane in a worksheet. This width is
	 * used only in the XML documents of worksheets and is not exposed to the
	 * (Excel) end user
	 *
	 * @param width
	 *            Target column(s) width (one or more columns, displayed in Excel)
	 * @return The internal pane width, used in worksheet XML documents in case of
	 *         worksheet splitting
	 * @apiNote The internal split width is based on the width of one or more
	 *          columns. It also depends on the maximum digit width of the default
	 *          font, as well as its text padding and various constants.<br>
	 *          See also {@link #getInternalColumnWidth(float, float, float)} for
	 *          additional details.<br>
	 *          This method is derived from the Perl implementation by John McNamara
	 *          (<a href=
	 *          "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br>
	 *          See also: <a href=
	 *          "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376,
	 *          Part 1, Chapter 18.3.1.13</a><br>
	 */
	public static float getInternalPaneSplitWidth(float width) {
		return getInternalPaneSplitWidth(width, 7f, 5f);
	}

	/**
	 * Calculates the internal width of a split pane in a worksheet. This width is
	 * used only in the XML documents of worksheets and is not exposed to the
	 * (Excel) end user
	 *
	 * @param width
	 *            Target column(s) width (one or more columns, displayed in Excel)
	 * @param maxDigitWidth
	 *            Maximum digit with of the default font (default is 7.0 for
	 *            Calibri, size 11)
	 * @param textPadding
	 *            Text padding of the default font (default is 5.0 for Calibri, size
	 *            11)
	 * @return The internal pane width, used in worksheet XML documents in case of
	 *         worksheet splitting
	 * @apiNote The internal split width is based on the width of one or more
	 *          columns. It also depends on the maximum digit width of the default
	 *          font, as well as its text padding and various constants.<br>
	 *          See also {@link #getInternalColumnWidth(float, float, float)} for
	 *          additional details.<br>
	 *          This method is derived from the Perl implementation by John McNamara
	 *          (<a href=
	 *          "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br>
	 *          See also: <a href=
	 *          "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376,
	 *          Part 1, Chapter 18.3.1.13</a><br>
	 *          The two override parameters maxDigitWidth and textPadding probably
	 *          don't have to be other than maxDigitWidth = 7f and textPadding = 5f.
	 */
	public static float getInternalPaneSplitWidth(float width, float maxDigitWidth, float textPadding) {
		float pixels;
		if (width < 0) {
			width = 0;
		}
		if (width <= 1f) {
			pixels = (float) Math.floor(width / SPLIT_WIDTH_MULTIPLIER + SPLIT_WIDTH_OFFSET);
		}
		else {
			pixels = (float) Math.floor(width * maxDigitWidth + SPLIT_WIDTH_OFFSET) + textPadding;
		}
		float points = pixels * SPLIT_WIDTH_POINT_MULTIPLIER;
		return points * SPLIT_POINT_DIVIDER + SPLIT_WIDTH_POINT_OFFSET;
	}

	/**
	 * Calculates the internal height of a split pane in a worksheet. This height is
	 * used only in the XML documents of worksheets and is not exposed to the
	 * (Excel) user
	 *
	 * @param height
	 *            Target row(s) height (one or more rows, displayed in Excel)
	 * @return The internal pane height, used in worksheet XML documents in case of
	 *         worksheet splitting
	 * @apiNote The internal split height is based on the height of one or more
	 *          rows. It also depends on various constants.<br>
	 *          This method is derived from the Perl implementation by John McNamara
	 *          (<a href=
	 *          "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>).<br>
	 *          Negative row heights are automatically transformed to 0.
	 */
	public static float getInternalPaneSplitHeight(float height) {
		if (height < 0) {
			height = 0f;
		}
		return (float) Math.floor(SPLIT_POINT_DIVIDER * height + SPLIT_HEIGHT_POINT_OFFSET);
	}

	/**
	 * Calculates the height of a split pane in a worksheet, based on the internal
	 * value (calculated by {@link Helper#getInternalPaneSplitHeight(float)})
	 *
	 * @param internalHeight
	 *            Internal pane height stored in a worksheet. The minimal value is
	 *            defined by {@link Helper#SPLIT_HEIGHT_POINT_OFFSET}
	 * @return Actual pane height
	 * @apiNote Depending on the initial height, the result value of
	 *          {@link Helper#getInternalPaneSplitHeight(float)} may not lead back
	 *          to the initial value, since rounding is applied when calculating the
	 *          internal height
	 */
	public static float getPaneSplitHeight(float internalHeight) {
		if (internalHeight < 300f) {
			return 0;
		}
		else {
			return (internalHeight - SPLIT_HEIGHT_POINT_OFFSET) / SPLIT_POINT_DIVIDER;
		}
	}

	/**
	 * Calculates the width of a split pane in a worksheet, based on the internal
	 * value (calculated by
	 * {@link Helper#getInternalPaneSplitWidth(float, float, float)})
	 *
	 * @param internalWidth
	 *            Internal pane width stored in a worksheet. The minimal value is
	 *            defined by {@link Helper#SPLIT_WIDTH_POINT_OFFSET}
	 * @return Actual pane width
	 * @apiNote Depending on the initial width, the result value of
	 *          {@link Helper#getInternalPaneSplitWidth(float)} or
	 *          {@link Helper#getInternalPaneSplitWidth(float, float, float)} may
	 *          not lead back to the initial value, since rounding is applied when
	 *          calculating the internal width
	 */
	public static float getPaneSplitWidth(float internalWidth) {
		return getPaneSplitWidth(internalWidth, 7f, 5f);
	}

	/**
	 * Calculates the width of a split pane in a worksheet, based on the internal
	 * value (calculated by
	 * {@link Helper#getInternalPaneSplitWidth(float, float, float)})
	 *
	 * @param internalWidth
	 *            Internal pane width stored in a worksheet. The minimal value is
	 *            defined by {@link Helper#SPLIT_WIDTH_POINT_OFFSET}
	 * @param maxDigitWidth
	 *            Maximum digit with of the default font (default is 7.0 for
	 *            Calibri, size 11)
	 * @param textPadding
	 *            Text padding of the default font (default is 5.0 for Calibri, size
	 *            11)
	 * @return Actual pane width
	 * @apiNote Depending on the initial width, the result value of
	 *          {@link Helper#getInternalPaneSplitWidth(float)} or
	 *          {@link Helper#getInternalPaneSplitWidth(float, float, float)} may
	 *          not lead back to the initial value, since rounding is applied when
	 *          calculating the internal width
	 */
	public static float getPaneSplitWidth(float internalWidth, float maxDigitWidth, float textPadding) {
		float points = (internalWidth - SPLIT_WIDTH_POINT_OFFSET) / SPLIT_POINT_DIVIDER;
		if (points < 0.001f) {
			return 0;
		}
		else {
			float width = points / SPLIT_WIDTH_POINT_MULTIPLIER;
			return (width - textPadding - SPLIT_WIDTH_OFFSET) / maxDigitWidth;
		}
	}

	/**
	 * Method to generate an Excel internal password hash to protect workbooks or
	 * worksheets<br>
	 * This method is derived from the c++ implementation by Kohei Yoshida (<a href=
	 * "http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)<br>
	 * <b>WARNING!</b> Do not use this method to encrypt 'real' passwords or data
	 * outside from PicoXLSX4j. This is only a minor security feature. Use a proper
	 * cryptography method instead.
	 *
	 * @param password
	 *            Password as plain text
	 * @return Encoded password
	 */
	public static String generatePasswordHash(String password) {
		if (Helper.isNullOrEmpty(password)) {
			return "";
		}
		int passwordLength = password.length();
		int passwordHash = 0;
		char character;
		for (int i = passwordLength; i > 0; i--) {
			character = password.charAt(i - 1);
			passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
			passwordHash ^= character;
		}
		passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
		passwordHash ^= (0x8000 | ('N' << 8) | 'K');
		passwordHash ^= passwordLength;
		return Integer.toHexString(passwordHash).toUpperCase();
	}

	/**
	 * Method to check a string whether its reference is null or the content is
	 * empty
	 *
	 * @param value
	 *            value / reference to check
	 * @return True if the passed parameter is null or empty, otherwise false
	 */
	public static boolean isNullOrEmpty(String value) {
		if (value == null) {
			return true;
		}
		return value.isEmpty();
	}

	/**
	 * Method to create a {@link Duration} object form hours, minutes and second
	 *
	 * @param hours
	 *            Number of Hours within the day
	 * @param minutes
	 *            Number of minutes within the hour
	 * @param seconds
	 *            Number of seconds within the minute
	 * @return Duration object
	 * @throws FormatException
	 *             thrown if one of the components is invalid (e.g. negative), or if
	 *             the number of days exceeds the year 9999
	 */
	public static Duration createDuration(int hours, int minutes, int seconds) {
		return createDuration(0, hours, minutes, seconds);
	}

	/**
	 * Method to create a {@link Duration} object form days, hours, minutes and
	 * second
	 *
	 * @param days
	 *            Total number of days
	 * @param hours
	 *            Number of Hours within the day
	 * @param minutes
	 *            Number of minutes within the hour
	 * @param seconds
	 *            Number of seconds within the minute
	 * @return Duration object
	 * @throws FormatException
	 *             thrown if one of the components is invalid (e.g. negative), or if
	 *             the number of days exceeds the year 9999
	 */
	public static Duration createDuration(int days, int hours, int minutes, int seconds) {
		if (days > MAX_OADATE_VALUE) {
			throw new FormatException("The number of days '" + days + "' exceeds the maximum allowed date of 9999-12-31");
		}
		if (days < 0 || hours < 0 || hours > 23 || minutes < 0 || minutes > 59 || seconds < 0 || seconds > 59) {
			throw new FormatException(String
					.format("One of the passed duration components is invalid. Days:%d, Hours:%d, Minutes:%d, Seconds:%d", days, hours, minutes, seconds));
		}
		Duration duration = Duration.ofDays(days);
		duration = duration.plusHours(hours).plusMinutes(minutes).plusSeconds(seconds);
		return duration;
	}

}
