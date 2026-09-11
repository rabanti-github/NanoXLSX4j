package ch.rabanti.nanoxlsx4j.utils;

import java.time.Duration;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class DataUtils {

    private DataUtils() {
        // Do not instantiate
    }

    /**
     * First date that can be displayed by Excel. Real values before this date cannot be processed.
     */
    public static final LocalDateTime FIRST_ALLOWED_EXCEL_DATE = LocalDateTime.of(1900, 1, 1, 0, 0, 0);

    /**
     * Last date that can be displayed by Excel. Real values after this date cannot be processed.
     */
    public static final LocalDateTime LAST_ALLOWED_EXCEL_DATE = LocalDateTime.of(9999, 12, 31, 23, 59, 59);

    /**
     * All dates before this date are shifted in Excel by -1.0, since Excel assumes wrongly that the year 1900 is a leap
     * year.<br /> See also: <a
     * href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     */
    public static final LocalDateTime FIRST_VALID_EXCEL_DATE = LocalDateTime.of(1900, 3, 1, 0, 0, 0);

    /**
     * Default time zone used for date conversion
     */
    public static final ZoneOffset DEFAULT_TIME_ZONE = ZoneOffset.UTC;

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
     * Strategy how ranges should be merged
     */
    public enum RangeMergeStrategy {
        /**
         * No merge should be performed
         */
        NO_MERGE,
        /**
         * Ranges of the same columns should be merged
         */
        MERGE_COLUMNS,
        /**
         * Ranges of the same row should be merged
         */
        MERGE_ROWS
    }

    /**
     * Method to convert a date or date and time into the internal Excel time format (OAdate)
     *
     * <p>Remarks: Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28
     * and
     * 1900-03-01. This method corrects all dates from the first valid date (1900-01-01) to 1900-03-01. However, Excel
     * displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month. In
     * conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999
     * milliseconds).<br /> See also: <a
     * href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br /> See also: <a
     * href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param date Date to process (time zone will be treated as UTC)
     * @return Date or date and time as number string
     * @throws FormatException Thrown if the passed date cannot be translated to the OADate format
     */
    public static String getOADateTimeString(Date date) {
        double d = getOADateTime(date);
        return ParserUtils.toString(d);
    }

    /**
     * Method to convert a date or date and time into the internal Excel time format (OAdate)
     *
     * <p>Remarks: Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28
     * and
     * 1900-03-01. This method corrects all dates from the first valid date (1900-01-01) to 1900-03-01. However, Excel
     * displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month. In
     * conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999
     * milliseconds).<br /> See also: <a
     * href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br /> See also: <a
     * href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param date Date to process (time zone will be treated as UTC)
     * @return Date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static double getOADateTime(Date date) {
        return getOADateTime(date, false);
    }

    /**
     * Method to convert a date or date and time into the internal Excel time format (OAdate)
     *
     * <p>Remarks: Excel assumes wrongly that the year 1900 is a leap year. There is a gap of 1.0 between 1900-02-28
     * and
     * 1900-03-01. This method corrects all dates from the first valid date (1900-01-01) to 1900-03-01. However, Excel
     * displays the minimum valid date as 1900-01-00, although 0 is not a valid description for a day of month. In
     * conformance to the OAdate specifications, the maximum valid date is 9999-12-31 23:59:59 (plus 999
     * milliseconds).<br /> See also: <a
     * href="https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1">
     * https://docs.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=netcore-3.1</a><br /> See also: <a
     * href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param skipCheck Flag to skip the validity check if set to true
     * @param date      Date to process (time zone will be treated as UTC)
     * @return Date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static double getOADateTime(Date date, boolean skipCheck) {
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
        if (dateValue.isBefore(FIRST_VALID_EXCEL_DATE)) {
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
     * <p>Remarks: The time is represented by a OAdate without the date component but a possible number of total
     * days</p>
     *
     * @param time Time to process. The date component of the timespan is converted to the total numbers of days
     * @return Time as number string
     */
    public static String getOATimeString(Duration time) {
        double d = getOATime(time);
        return ParserUtils.toString(d);
    }

    /// <summary>
    /// Method to convert a time into the internal Excel time format (OAdate without days)
    /// </summary>
    /// <param name="time">Time to process. The date component of the timespan is converted to the total numbers of
    /// days</param>
    /// <returns>Time as number</returns>
    /// \remark <remarks>The time is represented by a OAdate without the date component but a possible number of total
    /// days</remarks>
    public static double getOATime(Duration time) {
        return time.toSeconds() / 86400d;
    }

    /**
     * Method to calculate a common Date from the OA date (OLE automation) format<br /> OA Date format starts at January
     * 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and
     * will throw a FormatException
     *
     * <p>Remarks: Numbers that represents dates before 1900-03-01 (number of days since 1900-01-01 = 60) are
     * automatically modified.
     * Until 1900-03-01 is 1.0 added to the number to get the same date, as displayed in Excel.The reason for this is a
     * bug in Excel. See also: <a
     * href="https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year">
     * https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year</a>
     * </p>
     *
     * @param oaDate oaDate OA date number
     * @return Converted date
     */
    public static Date getDateFromOA(double oaDate) {
        if (oaDate < 60d) {
            oaDate++;
        }
        long milliseconds = Math.round(oaDate * 86400000d);
        return Date.from(ROOT_DATE.toInstant(DEFAULT_TIME_ZONE).plusMillis(milliseconds));
    }

    /**
     * Calculates the internal width of a column in characters. This width is used only in the XML documents of
     * worksheets and is usually not exposed to the (Excel) end user
     *
     * @param columnWidth Target column width (displayed in Excel)
     * @return The internal column width in characters, used in worksheet XML documents
     * @apiNote The internal width deviates slightly from the column width, entered in Excel. Although internal, the
     * default column width of 10 characters is visible in Excel as 10.71. The deviation depends on the maximum digit
     * width of the default font, as well as its text padding and various constants.<br> In case of the width 10.0 and
     * the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11), the internal width
     * is approximately 10.7142857 (rounded to 10.71).<br> Note that the column height is not affected by this
     * consideration. The entered height in Excel is the actual height in the worksheet XML documents.<br> This method
     * is derived from the Perl implementation by John McNamara (<a href=
     * "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br> See also: <a href=
     * "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter
     * 18.3.1.13</a>
     */
    public static float getInternalColumnWidth(float columnWidth) {
        return getInternalColumnWidth(columnWidth, 7f, 5f);
    }

    /**
     * Calculates the internal width of a column in characters. This width is used only in the XML documents of
     * worksheets and is usually not exposed to the (Excel) end user
     *
     * @param columnWidth   Target column width (displayed in Excel)
     * @param maxDigitWidth Maximum digit with of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return The internal column width in characters, used in worksheet XML documents
     * @throws FormatException thrown if the column width is out of range
     * @apiNote The internal width deviates slightly from the column width, entered in Excel. Although internal, the
     * default column width of 10 characters is visible in Excel as 10.71. The deviation depends on the maximum digit
     * width of the default font, as well as its text padding and various constants.<br> In case of the width 10.0 and
     * the default digit width 7.0, as well as the padding 5.0 of the default font Calibri (size 11), the internal width
     * is approximately 10.7142857 (rounded to 10.71).<br> Note that the column height is not affected by this
     * consideration. The entered height in Excel is the actual height in the worksheet XML documents.<br> This method
     * is derived from the Perl implementation by John McNamara (<a href=
     * "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br> See also: <a href=
     * "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter
     * 18.3.1.13</a>
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
        } else if (columnWidth <= 1f) {
            return (float) Math.floor(
                    (columnWidth * (maxDigitWidth + textPadding)) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) /
                    COLUMN_WIDTH_ROUNDING_MODIFIER;
        } else {
            return (float) Math.floor(
                    (columnWidth * maxDigitWidth + textPadding) / maxDigitWidth * COLUMN_WIDTH_ROUNDING_MODIFIER) /
                    COLUMN_WIDTH_ROUNDING_MODIFIER;
        }
    }

    /**
     * Calculates the internal height of a row. This height is used only in the XML documents of worksheets and is
     * usually not exposed to the (Excel) end user
     *
     * @param rowHeight Target row height (displayed in Excel)
     * @return The internal row height which snaps to the nearest pixel
     * @throws FormatException thrown if the row height is out of range
     * @apiNote The height is based on the calculated amount of pixels. One point are ~1.333 (1+1/3) pixels. After the
     * conversion, the number of pixels is rounded to the nearest integer and calculated back to points.<br> Therefore,
     * the originally defined row height will slightly deviate, based on this pixel snap
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
     * Calculates the internal width of a split pane in a worksheet. This width is used only in the XML documents of
     * worksheets and is not exposed to the (Excel) end user
     *
     * @param width Target column(s) width (one or more columns, displayed in Excel)
     * @return The internal pane width, used in worksheet XML documents in case of worksheet splitting
     * @apiNote The internal split width is based on the width of one or more columns. It also depends on the maximum
     * digit width of the default font, as well as its text padding and various constants.<br> See also
     * {@link #getInternalColumnWidth(float, float, float)} for additional details.<br> This method is derived from the
     * Perl implementation by John McNamara (<a href=
     * "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br> See also: <a href=
     * "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter
     * 18.3.1.13</a><br>
     */
    public static float getInternalPaneSplitWidth(float width) {
        return getInternalPaneSplitWidth(width, 7f, 5f);
    }

    /**
     * Calculates the internal width of a split pane in a worksheet. This width is used only in the XML documents of
     * worksheets and is not exposed to the (Excel) end user
     *
     * @param width         Target column(s) width (one or more columns, displayed in Excel)
     * @param maxDigitWidth Maximum digit with of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return The internal pane width, used in worksheet XML documents in case of worksheet splitting
     * @apiNote The internal split width is based on the width of one or more columns. It also depends on the maximum
     * digit width of the default font, as well as its text padding and various constants.<br> See also
     * {@link #getInternalColumnWidth(float, float, float)} for additional details.<br> This method is derived from the
     * Perl implementation by John McNamara (<a href=
     * "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>)<br> See also: <a href=
     * "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/">ECMA-376, Part 1, Chapter
     * 18.3.1.13</a><br> The two override parameters maxDigitWidth and textPadding probably don't have to be other than
     * maxDigitWidth = 7f and textPadding = 5f.
     */
    public static float getInternalPaneSplitWidth(float width, float maxDigitWidth, float textPadding) {
        float pixels;
        if (width < 0) {
            width = 0;
        }
        if (width <= 1f) {
            pixels = (float) Math.floor(width / SPLIT_WIDTH_MULTIPLIER + SPLIT_WIDTH_OFFSET);
        } else {
            pixels = (float) Math.floor(width * maxDigitWidth + SPLIT_WIDTH_OFFSET) + textPadding;
        }
        float points = pixels * SPLIT_WIDTH_POINT_MULTIPLIER;
        return points * SPLIT_POINT_DIVIDER + SPLIT_WIDTH_POINT_OFFSET;
    }

    /**
     * Calculates the internal height of a split pane in a worksheet. This height is used only in the XML documents of
     * worksheets and is not exposed to the (Excel) user
     *
     * @param height Target row(s) height (one or more rows, displayed in Excel)
     * @return The internal pane height, used in worksheet XML documents in case of worksheet splitting
     * @apiNote The internal split height is based on the height of one or more rows. It also depends on various
     * constants.<br> This method is derived from the Perl implementation by John McNamara (<a href=
     * "https://stackoverflow.com/a/5010899">https://stackoverflow.com/a/5010899</a>).<br> Negative row heights are
     * automatically transformed to 0.
     */
    public static float getInternalPaneSplitHeight(float height) {
        if (height < 0) {
            height = 0f;
        }
        return (float) Math.floor(SPLIT_POINT_DIVIDER * height + SPLIT_HEIGHT_POINT_OFFSET);
    }

    /**
     * Calculates the height of a split pane in a worksheet, based on the internal value (calculated by
     * {@link DataUtils#getInternalPaneSplitHeight(float)})
     *
     * @param internalHeight Internal pane height stored in a worksheet. The minimal value is defined by
     *                       {@link DataUtils#SPLIT_HEIGHT_POINT_OFFSET}
     * @return Actual pane height
     * @apiNote Depending on the initial height, the result value of {@link DataUtils#getInternalPaneSplitHeight(float)}
     * may not lead back to the initial value, since rounding is applied when calculating the internal height
     */
    public static float getPaneSplitHeight(float internalHeight) {
        if (internalHeight < 300f) {
            return 0;
        } else {
            return (internalHeight - SPLIT_HEIGHT_POINT_OFFSET) / SPLIT_POINT_DIVIDER;
        }
    }

    /**
     * Calculates the width of a split pane in a worksheet, based on the internal value (calculated by
     * {@link DataUtils#getInternalPaneSplitWidth(float, float, float)})
     *
     * @param internalWidth Internal pane width stored in a worksheet. The minimal value is defined by
     *                      {@link DataUtils#SPLIT_WIDTH_POINT_OFFSET}
     * @return Actual pane width
     * @apiNote Depending on the initial width, the result value of {@link DataUtils#getInternalPaneSplitWidth(float)}
     * or {@link DataUtils#getInternalPaneSplitWidth(float, float, float)} may not lead back to the initial value, since
     * rounding is applied when calculating the internal width
     */
    public static float getPaneSplitWidth(float internalWidth) {
        return getPaneSplitWidth(internalWidth, 7f, 5f);
    }

    /**
     * Calculates the width of a split pane in a worksheet, based on the internal value (calculated by
     * {@link DataUtils#getInternalPaneSplitWidth(float, float, float)})
     *
     * @param internalWidth Internal pane width stored in a worksheet. The minimal value is defined by
     *                      {@link DataUtils#SPLIT_WIDTH_POINT_OFFSET}
     * @param maxDigitWidth Maximum digit with of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return Actual pane width
     * @apiNote Depending on the initial width, the result value of {@link DataUtils#getInternalPaneSplitWidth(float)}
     * or {@link DataUtils#getInternalPaneSplitWidth(float, float, float)} may not lead back to the initial value, since
     * rounding is applied when calculating the internal width
     */
    public static float getPaneSplitWidth(float internalWidth, float maxDigitWidth, float textPadding) {
        float points = (internalWidth - SPLIT_WIDTH_POINT_OFFSET) / SPLIT_POINT_DIVIDER;
        if (points < 0.001f) {
            return 0;
        } else {
            float width = points / SPLIT_WIDTH_POINT_MULTIPLIER;
            return (width - textPadding - SPLIT_WIDTH_OFFSET) / maxDigitWidth;
        }
    }

    /**
     * Merges a range with a list of given ranges. If there is no intersection between the list and the new range, the
     * range is just added to the given list. If there is an intersection, the range will be merged and the new list of
     * ranges will be returned. The strategy is {@link RangeMergeStrategy#MERGE_COLUMNS}.
     *
     * @param givenRanges List of given ranges
     * @param newRange    The range to be merged
     * @return List of resulting ranges after merging.
     */
    public static List<Range> mergeRange(List<Range> givenRanges, Range newRange) {
        return mergeRange(givenRanges, newRange, RangeMergeStrategy.MERGE_COLUMNS);
    }

    /**
     * Merges a range with a list of given ranges. If there is no intersection between the list and the new range, the
     * range is just added to the given list. If there is an intersection, the range will be merged and the new list of
     * ranges will be returned
     *
     * @param givenRanges List of given ranges
     * @param newRange    The range to be merged
     * @param strategy    Strategy for the range recalculation. Depending on the value, the resulting ranges are either
     *                    merged along rows, along columns (default), or not merged at all
     * @return List of resulting ranges after merging.
     */
    public static List<Range> mergeRange(List<Range> givenRanges, Range newRange, RangeMergeStrategy strategy) {
        List<Range> result = new ArrayList<>();
        List<Range> mergedCandidates = new ArrayList<>();
        mergedCandidates.add(newRange);
        // Step 1: Find intersecting ranges and remove them from existingRanges
        for (Range range : givenRanges) {
            if (isMergeCandidate(newRange, range, strategy)) {
                mergedCandidates.add(range);
            } else {
                result.add(range);
            }
        }
        // Step 2: Slice intersecting/adjacent ranges into uniform rectangular pieces.
        List<Range> slicedRanges = sliceRanges(mergedCandidates);
        // Step 3: Merge adjacent rectangles where possible.
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            result.addAll(mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_COLUMNS));
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_ROWS);
        } else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            result.addAll(mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_ROWS));
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_COLUMNS);
        } else {
            result.addAll(slicedRanges);
        }
        return result;
    }

    /**
     * Returns true if the two ranges are either overlapping or adjacent in the appropriate direction for the chosen
     * merge strategy. For vertical merging (MergeColumns): the ranges must share the same column boundaries and be
     * either overlapping or immediately adjacent vertically. For horizontal merging (MergeRows): the ranges must share
     * the same row boundaries and be either overlapping or immediately adjacent horizontally.
     */
    private static boolean isMergeCandidate(Range a, Range b, RangeMergeStrategy strategy) {
        // First, if they overlap, they are candidates.
        if (a.overlaps(b)) {
            return true;
        }

        // Otherwise, check for adjacency according to the strategy.
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            // Vertical merging: require same columns.
            if (a.startAddress().column() == b.startAddress().column() &&
                    a.endAddress().column() == b.endAddress().column() &&
                    (a.endAddress().row() + 1 == b.startAddress().row() ||
                            b.endAddress().row() + 1 == a.startAddress().row())) {
                return true;
            }
        } else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            // Horizontal merging: require same rows.
            if (a.startAddress().row() == b.startAddress().row() &&
                    a.endAddress().row() == b.endAddress().row() &&
                    (a.endAddress().column() + 1 == b.startAddress().column() ||
                            b.endAddress().column() + 1 == a.startAddress().column())) {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Subtracts a range form a list of given ranges. If the range to be removed does not intersect any of the given
    /// ranges, nothing happens. If the range intersects at least one of the given ranges, the intersection will be
    /// removed and the new ranges well be returned. The strategy is {@link RangeMergeStrategy#MERGE_COLUMNS}.
    /// </summary>
    /// <param name="givenRanges">List of given ranges</param>
    /// <param name="rangeToRemove">The range to be removed</param>
    /// <param name="strategy">Strategy for the range recalculation. Depending on the value, the resulting ranges are
    /// either merged along rows, along columns (default), or not merged at all</param>
    /// <returns>List of resulting ranges after subtraction and recalculation</returns>
    public static List<Range> subtractRange(List<Range> givenRanges, Range rangeToRemove) {
        return subtractRange(givenRanges, rangeToRemove, RangeMergeStrategy.MERGE_COLUMNS);
    }

    /// <summary>
    /// Subtracts a range form a list of given ranges. If the range to be removed does not intersect any of the given
    /// ranges, nothing happens. If the range intersects at least one of the given ranges, the intersection will be
    /// removed and the new ranges well be returned.
    /// </summary>
    /// <param name="givenRanges">List of given ranges</param>
    /// <param name="rangeToRemove">The range to be removed</param>
    /// <param name="strategy">Strategy for the range recalculation. Depending on the value, the resulting ranges are
    /// either merged along rows, along columns (default), or not merged at all</param>
    /// <returns>List of resulting ranges after subtraction and recalculation</returns>
    public static List<Range> subtractRange(List<Range> givenRanges, Range rangeToRemove, RangeMergeStrategy strategy) {
        List<Range> result = new ArrayList<>();
        // Process each existing range.
        for (Range range : givenRanges) {
            if (!range.overlaps(rangeToRemove)) {
                // No overlap: keep the range unchanged.
                result.add(range);
            } else {
                // Overlapping range: subtract the removal area.
                List<Range> subtractedPieces = subtractRect(range, rangeToRemove);
                result.addAll(subtractedPieces);
            }
        }
        // Slice all ranges before merge
        List<Range> slicedRanges = sliceRanges(result);
        // Merge adjacent pieces if requested.
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            result = mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_COLUMNS);
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_ROWS);
        } else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            result = mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_ROWS);
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_COLUMNS);
        } else {
            result = slicedRanges;
        }
        return result;
    }

    /**
     * Method to slice possibly overlapping ranges into contiguous ranges without intersections
     *
     * @param ranges Ranges to slice
     * @return List of sliced, contiguous ranges
     */
    private static List<Range> sliceRanges(List<Range> ranges) {
        Set<Integer> uniqueCols = new HashSet<>();
        Set<Integer> uniqueRows = new HashSet<>();

        // Collect all column and row boundaries
        for (Range range : ranges) {
            uniqueCols.add(range.startAddress().column());
            uniqueCols.add(range.endAddress().column() + 1); // To handle gaps properly
            uniqueRows.add(range.startAddress().row());
            uniqueRows.add(range.endAddress().row() + 1);
        }

        // Convert to sorted lists for iteration
        List<Integer> sortedCols = uniqueCols.stream().sorted().toList();
        List<Integer> sortedRows = uniqueRows.stream().sorted().toList();

        List<Range> slicedRanges = new ArrayList<>();

        // Step through the row and column boundaries to create the smallest sub-rectangles
        for (int r = 0; r < sortedRows.size() - 1; r++) {
            for (int c = 0; c < sortedCols.size() - 1; c++) {
                Range subRange = new Range(
                        sortedCols.get(c),
                        sortedRows.get(r),
                        sortedCols.get(c + 1) - 1,
                        sortedRows.get(r + 1) - 1
                );

                // Only keep the sub-range if it was originally covered
                if (ranges.stream().anyMatch(range -> range.contains(subRange))) {
                    slicedRanges.add(subRange);
                }
            }
        }
        return slicedRanges;
    }

    /**
     * Subtracts the removal range from an original range. Returns up to 4 rectangular pieces that cover (original minus
     * the intersecting part). If there is no intersection, returns the original range.
     */
    private static List<Range> subtractRect(Range original, Range toRemove) {
        List<Range> pieces = new ArrayList<>();
        // Original boundaries:
        int orig_left = original.startAddress().column();
        int orig_top = original.startAddress().row();
        int orig_right = original.endAddress().column();
        int orig_bottom = original.endAddress().row();
        // Removal boundaries:
        int rem_left = toRemove.startAddress().column();
        int rem_top = toRemove.startAddress().row();
        int rem_right = toRemove.endAddress().column();
        int rem_bottom = toRemove.endAddress().row();
        // Compute intersection boundaries.
        int isct_left = Math.max(orig_left, rem_left);
        int isct_top = Math.max(orig_top, rem_top);
        int isct_right = Math.min(orig_right, rem_right);
        int isct_bottom = Math.min(orig_bottom, rem_bottom);

        // Slice the original rectangle into up to four pieces.
        // Top piece: if any rows exist above the intersection.
        if (orig_top < isct_top) {
            pieces.add(new Range(orig_left, orig_top, orig_right, isct_top - 1));
        }
        // Bottom piece: if any rows exist below the intersection.
        if (isct_bottom < orig_bottom) {
            pieces.add(new Range(orig_left, isct_bottom + 1, orig_right, orig_bottom));
        }
        // Left piece: if any columns exist to the left of the intersection within the vertical boundaries of the
        // intersection.
        if (orig_left < isct_left) {
            pieces.add(new Range(orig_left, isct_top, isct_left - 1, isct_bottom));
        }
        // Right piece: if any columns exist to the right of the intersection within the vertical boundaries of the
        // intersection.
        if (isct_right < orig_right) {
            pieces.add(new Range(isct_right + 1, isct_top, orig_right, isct_bottom));
        }
        return pieces;
    }

    /**
     * Method to merge ranges by rows or columns (according to the strategy) that can be merged into a new ranges, so
     * that all addresses of the original ranges are still covered and no additional addresses are used.
     *
     * @param ranges   Original (sliced) ranges
     * @param strategy Merge strategy. If the strategy is NoMerge, the original list will be returned
     * @return List of merged ranges
     */
    private static List<Range> mergeAdjacentRanges(List<Range> ranges, RangeMergeStrategy strategy) {
        if (ranges.isEmpty()) {
            return new ArrayList<>();
        }
        List<Range> mergedRanges = new ArrayList<>();

        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            // Vertical merging: Ranges must have identical column boundaries.
            // Group by StartAddress.Column and EndAddress.Column.
            record ColumnBounds(int startCol, int endCol) {
            }
            var groups = ranges.stream()
                    .collect(Collectors.groupingBy(r ->
                            new ColumnBounds(
                                    r.startAddress().column(),
                                    r.endAddress().column()
                            )
                    ));
            for (var group : groups.entrySet()) {
                // Order by row (ascending)
                List<Range> sorted = group.getValue().stream()
                        .sorted(Comparator.comparingInt(r -> r.startAddress().row()))
                        .toList();
                Range current = sorted.getFirst();
                for (int i = 1; i < sorted.size(); i++) {
                    Range next = sorted.get(i);
                    // Check if the current range is contiguous with or overlapping the next.
                    // (That is, if current.EndAddress.Row + 1 is >= next.StartAddress.Row.)
                    if (current.endAddress().row() + 1 >= next.startAddress().row()) {
                        // They share the same columns.
                        // Create a new range from current.StartAddress.Row to the maximum of the two EndAddress.Row
                        // values.
                        int newStartRow = current.startAddress().row();
                        int newEndRow = Math.max(current.endAddress().row(), next.endAddress().row());
                        current = new Range(
                                current.startAddress().column(), newStartRow,
                                current.endAddress().column(), newEndRow
                        );
                    } else {
                        mergedRanges.add(current);
                        current = next;
                    }
                }
                mergedRanges.add(current);
            }
        } else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            // Horizontal merging: Ranges must have identical row boundaries.
            // Group by StartAddress.Row and EndAddress.Row.
            record RowBounds(int startRow, int endRow) {
            }
            var groups = ranges.stream()
                    .collect(Collectors.groupingBy(r ->
                            new RowBounds(
                                    r.startAddress().row(),
                                    r.endAddress().row()
                            )
                    ));

            for (var group : groups.entrySet()) {
                List<Range> sorted = group.getValue().stream()
                        .sorted(Comparator.comparingInt(r -> r.startAddress().column()))
                        .toList();
                Range current = sorted.getFirst();
                for (int i = 1; i < sorted.size(); i++) {
                    Range next = sorted.get(i);
                    // Check if current.EndAddress.Column + 1 is >= next.StartAddress.Column.
                    if (current.endAddress().column() + 1 >= next.startAddress().column()) {
                        int newStartCol = current.startAddress().column();
                        int newEndCol = Math.max(current.endAddress().column(), next.endAddress().column());
                        current = new Range(
                                newStartCol, current.startAddress().row(),
                                newEndCol, current.endAddress().row()
                        );
                    } else {
                        mergedRanges.add(current);
                        current = next;
                    }
                }
                mergedRanges.add(current);
            }
        }
        return mergedRanges;
    }


}
