/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.utils;

import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

import java.time.Duration;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalAccessor;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * General data utils class with static methods
 *
 * @author Raphael Stoeckli
 */
public class DataUtils {

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
     * Minimum valid OAdate value (1900-01-01). However, Excel displays this value as 1900-01-00 (day zero)
     */
    public static final double MIN_OADATE_VALUE = 0;
    /**
     * Maximum valid OAdate value (9999-12-31)
     */
    public static final double MAX_OADATE_VALUE = 2958465.999988426d;
    /**
     * First date that can be displayed by Excel. Values before this date cannot be processed.
     */
    public static final Date FIRST_ALLOWED_EXCEL_DATE;

    /**
     * Last date that can be displayed by Excel. Values after this date cannot be processed.
     */
    public static final Date LAST_ALLOWED_EXCEL_DATE;

    /**
     * All dates before this date are shifted in Excel by -1.0, since Excel assumes wrongly that the year 1900 is a leap
     * year.<br> See also: <a href=
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

    // ### E N U M S ###

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
         * Ranges of the same rows should be merged
         */
        MERGE_ROWS
    }

    // ### C O N S T R U C T O R S ###
    private DataUtils() {
        // Prevents class instantiation
    }

    // ### S T A T I C M E T H O D S ###

    /**
     * Method to calculate the OA date (OLE automation) of the passed date.<br>
     *
     * @param date Date to convert
     * @return OA date or date and time as number string
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static String getOADateString(Date date) {
        double d = getOADate(date);
        return Double.toString(d);
    }

    /**
     * Method to calculate the OA date (OLE automation) of the passed date
     *
     * @param date Date to convert
     * @return OA date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     */
    public static double getOADate(Date date) {
        return getOADate(date, false);
    }

    /**
     * Method to calculate the OA date (OLE automation) of the passed date
     *
     * @param date      Date to convert
     * @param skipCheck If true, the validity check is skipped
     * @return OA date or date and time as number
     * @throws FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
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
     * Method to convert a duration into the internal Excel time format (OAdate without days)
     *
     * @param time Time to process
     * @return Time as number string
     * @throws FormatException Throws a FormatException if the passed time is invalid
     */
    public static String getOATimeString(Duration time) {
        double d = getOATime(time);
        return Double.toString(d);
    }

    /**
     * Method to convert a duration into the internal Excel time format (OAdate without days)
     *
     * @param time Time to process
     * @return Time as number
     * @throws FormatException Throws a FormatException if the passed time is invalid
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
     * Method to calculate a common Date from the OA date (OLE automation) format
     *
     * @param oaDate OA date number
     * @return Converted date
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
     * Method to parse a {@link Duration} object from a string and a pattern
     *
     * @param timeString String to parse
     * @param pattern    Pattern as string
     * @param locale     Locale of the formatter
     * @return Duration object
     * @throws FormatException thrown if the time could not be parsed or if the pattern was invalid
     */
    public static Duration parseTime(String timeString, String pattern, Locale locale) {
        if (ParserUtils.isNullOrEmpty(timeString) || ParserUtils.isNullOrEmpty(pattern)) {
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
     * Method to parse a {@link Duration} object from a string and a {@link DateTimeFormatter} instance
     *
     * @param timeString String to parse
     * @param formatter  Formatter to apply the parsing
     * @return Duration object
     * @throws FormatException thrown if the time could not be parsed or if the formatter was invalid
     */
    public static Duration parseTime(String timeString, DateTimeFormatter formatter) {
        if (ParserUtils.isNullOrEmpty(timeString) || formatter == null) {
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
     * Calculates the internal width of a column in characters
     *
     * @param columnWidth Target column width (displayed in Excel)
     * @return The internal column width in characters, used in worksheet XML documents
     * @throws FormatException thrown if the column width is out of range
     */
    public static float getInternalColumnWidth(float columnWidth) {
        return getInternalColumnWidth(columnWidth, 7f, 5f);
    }

    /**
     * Calculates the internal width of a column in characters
     *
     * @param columnWidth   Target column width (displayed in Excel)
     * @param maxDigitWidth Maximum digit width of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return The internal column width in characters, used in worksheet XML documents
     * @throws FormatException thrown if the column width is out of range
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
     * Calculates the internal height of a row
     *
     * @param rowHeight Target row height (displayed in Excel)
     * @return The internal row height which snaps to the nearest pixel
     * @throws FormatException thrown if the row height is out of range
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
     * Calculates the internal width of a split pane in a worksheet
     *
     * @param width Target column(s) width (one or more columns, displayed in Excel)
     * @return The internal pane width, used in worksheet XML documents in case of worksheet splitting
     */
    public static float getInternalPaneSplitWidth(float width) {
        return getInternalPaneSplitWidth(width, 7f, 5f);
    }

    /**
     * Calculates the internal width of a split pane in a worksheet
     *
     * @param width         Target column(s) width (one or more columns, displayed in Excel)
     * @param maxDigitWidth Maximum digit width of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return The internal pane width, used in worksheet XML documents in case of worksheet splitting
     */
    public static float getInternalPaneSplitWidth(float width, float maxDigitWidth, float textPadding) {
        float pixels;
        if (width <= 1f) {
            // NOTE: values <= 1 (including negative) are treated as zero and always result in 390
            width = 0;
            pixels = (float) Math.floor(width / SPLIT_WIDTH_MULTIPLIER + SPLIT_WIDTH_OFFSET);
        }
        else {
            pixels = (float) Math.floor(width * maxDigitWidth + SPLIT_WIDTH_OFFSET) + textPadding;
        }
        float points = pixels * SPLIT_WIDTH_POINT_MULTIPLIER;
        return points * SPLIT_POINT_DIVIDER + SPLIT_WIDTH_POINT_OFFSET;
    }

    /**
     * Calculates the internal height of a split pane in a worksheet
     *
     * @param height Target row(s) height (one or more rows, displayed in Excel)
     * @return The internal pane height, used in worksheet XML documents in case of worksheet splitting
     */
    public static float getInternalPaneSplitHeight(float height) {
        if (height < 0) {
            height = 0f;
        }
        return (float) Math.floor(SPLIT_POINT_DIVIDER * height + SPLIT_HEIGHT_POINT_OFFSET);
    }

    /**
     * Calculates the height of a split pane in a worksheet, based on the internal value
     *
     * @param internalHeight Internal pane height stored in a worksheet
     * @return Actual pane height
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
     * Calculates the width of a split pane in a worksheet, based on the internal value
     *
     * @param internalWidth Internal pane width stored in a worksheet
     * @return Actual pane width
     */
    public static float getPaneSplitWidth(float internalWidth) {
        return getPaneSplitWidth(internalWidth, 7f, 5f);
    }

    /**
     * Calculates the width of a split pane in a worksheet, based on the internal value
     *
     * @param internalWidth Internal pane width stored in a worksheet
     * @param maxDigitWidth Maximum digit width of the default font (default is 7.0 for Calibri, size 11)
     * @param textPadding   Text padding of the default font (default is 5.0 for Calibri, size 11)
     * @return Actual pane width
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
     * Merges a range with a list of given ranges using the default MERGE_COLUMNS strategy.
     * If there is no intersection between the list and the new range, the range is just added.
     * If there is an intersection, the ranges will be merged.
     *
     * @param givenRanges List of given ranges
     * @param newRange    The range to be merged
     * @return List of resulting ranges after merging
     */
    public static List<Range> mergeRange(List<Range> givenRanges, Range newRange) {
        return mergeRange(givenRanges, newRange, RangeMergeStrategy.MERGE_COLUMNS);
    }

    /**
     * Merges a range with a list of given ranges. If there is no intersection between the list and the new range,
     * the range is just added. If there is an intersection, the range will be merged.
     *
     * @param givenRanges List of given ranges
     * @param newRange    The range to be merged
     * @param strategy    Strategy for the range recalculation
     * @return List of resulting ranges after merging
     */
    public static List<Range> mergeRange(List<Range> givenRanges, Range newRange, RangeMergeStrategy strategy) {
        List<Range> result = new ArrayList<>();
        List<Range> mergedCandidates = new ArrayList<>();
        mergedCandidates.add(newRange);
        for (Range range : givenRanges) {
            if (isMergeCandidate(newRange, range, strategy)) {
                mergedCandidates.add(range);
            }
            else {
                result.add(range);
            }
        }
        List<Range> slicedRanges = sliceRanges(mergedCandidates);
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            result.addAll(mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_COLUMNS));
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_ROWS);
        }
        else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            result.addAll(mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_ROWS));
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_COLUMNS);
        }
        else {
            result.addAll(slicedRanges);
        }
        return result;
    }

    /**
     * Subtracts a range from a list of given ranges using the default MERGE_COLUMNS strategy.
     *
     * @param givenRanges   List of given ranges
     * @param rangeToRemove The range to be removed
     * @return List of resulting ranges after subtraction
     */
    public static List<Range> subtractRange(List<Range> givenRanges, Range rangeToRemove) {
        return subtractRange(givenRanges, rangeToRemove, RangeMergeStrategy.MERGE_COLUMNS);
    }

    /**
     * Subtracts a range from a list of given ranges. If the range to be removed does not intersect any of the
     * given ranges, nothing changes. If it intersects, the intersection will be removed.
     *
     * @param givenRanges   List of given ranges
     * @param rangeToRemove The range to be removed
     * @param strategy      Strategy for the range recalculation
     * @return List of resulting ranges after subtraction
     */
    public static List<Range> subtractRange(List<Range> givenRanges, Range rangeToRemove, RangeMergeStrategy strategy) {
        List<Range> result = new ArrayList<>();
        for (Range range : givenRanges) {
            if (!range.overlaps(rangeToRemove)) {
                result.add(range);
            }
            else {
                result.addAll(subtractRect(range, rangeToRemove));
            }
        }
        List<Range> slicedRanges = sliceRanges(result);
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            result = mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_COLUMNS);
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_ROWS);
        }
        else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            result = mergeAdjacentRanges(slicedRanges, RangeMergeStrategy.MERGE_ROWS);
            result = mergeAdjacentRanges(result, RangeMergeStrategy.MERGE_COLUMNS);
        }
        else {
            result = slicedRanges;
        }
        return result;
    }

    private static boolean isMergeCandidate(Range a, Range b, RangeMergeStrategy strategy) {
        if (a.overlaps(b)) {
            return true;
        }
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            if (a.getStartAddress().getColumn() == b.getStartAddress().getColumn() &&
                    a.getEndAddress().getColumn() == b.getEndAddress().getColumn() &&
                    (a.getEndAddress().getRow() + 1 == b.getStartAddress().getRow() ||
                            b.getEndAddress().getRow() + 1 == a.getStartAddress().getRow())) {
                return true;
            }
        }
        else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            if (a.getStartAddress().getRow() == b.getStartAddress().getRow() &&
                    a.getEndAddress().getRow() == b.getEndAddress().getRow() &&
                    (a.getEndAddress().getColumn() + 1 == b.getStartAddress().getColumn() ||
                            b.getEndAddress().getColumn() + 1 == a.getStartAddress().getColumn())) {
                return true;
            }
        }
        return false;
    }

    private static List<Range> sliceRanges(List<Range> ranges) {
        Set<Integer> uniqueCols = new HashSet<>();
        Set<Integer> uniqueRows = new HashSet<>();
        for (Range range : ranges) {
            uniqueCols.add(range.getStartAddress().getColumn());
            uniqueCols.add(range.getEndAddress().getColumn() + 1);
            uniqueRows.add(range.getStartAddress().getRow());
            uniqueRows.add(range.getEndAddress().getRow() + 1);
        }
        List<Integer> sortedCols = new ArrayList<>(uniqueCols);
        Collections.sort(sortedCols);
        List<Integer> sortedRows = new ArrayList<>(uniqueRows);
        Collections.sort(sortedRows);
        List<Range> slicedRanges = new ArrayList<>();
        for (int r = 0; r < sortedRows.size() - 1; r++) {
            for (int c = 0; c < sortedCols.size() - 1; c++) {
                Range subRange = new Range(sortedCols.get(c), sortedRows.get(r), sortedCols.get(c + 1) - 1, sortedRows.get(r + 1) - 1);
                boolean covered = ranges.stream().anyMatch(range -> range.contains(subRange));
                if (covered) {
                    slicedRanges.add(subRange);
                }
            }
        }
        return slicedRanges;
    }

    private static List<Range> subtractRect(Range original, Range toRemove) {
        List<Range> pieces = new ArrayList<>();
        int origLeft = original.getStartAddress().getColumn();
        int origTop = original.getStartAddress().getRow();
        int origRight = original.getEndAddress().getColumn();
        int origBottom = original.getEndAddress().getRow();
        int remLeft = toRemove.getStartAddress().getColumn();
        int remTop = toRemove.getStartAddress().getRow();
        int remRight = toRemove.getEndAddress().getColumn();
        int remBottom = toRemove.getEndAddress().getRow();
        int isctLeft = Math.max(origLeft, remLeft);
        int isctTop = Math.max(origTop, remTop);
        int isctRight = Math.min(origRight, remRight);
        int isctBottom = Math.min(origBottom, remBottom);
        if (origTop < isctTop) {
            pieces.add(new Range(origLeft, origTop, origRight, isctTop - 1));
        }
        if (isctBottom < origBottom) {
            pieces.add(new Range(origLeft, isctBottom + 1, origRight, origBottom));
        }
        if (origLeft < isctLeft) {
            pieces.add(new Range(origLeft, isctTop, isctLeft - 1, isctBottom));
        }
        if (isctRight < origRight) {
            pieces.add(new Range(isctRight + 1, isctTop, origRight, isctBottom));
        }
        return pieces;
    }

    private static List<Range> mergeAdjacentRanges(List<Range> ranges, RangeMergeStrategy strategy) {
        if (ranges.isEmpty()) {
            return new ArrayList<>();
        }
        List<Range> mergedRanges = new ArrayList<>();
        if (strategy == RangeMergeStrategy.MERGE_COLUMNS) {
            Map<String, List<Range>> groups = ranges.stream()
                    .collect(Collectors.groupingBy(r -> r.getStartAddress().getColumn() + "," + r.getEndAddress().getColumn()));
            for (List<Range> group : groups.values()) {
                List<Range> sorted = group.stream()
                        .sorted(Comparator.comparingInt(r -> r.getStartAddress().getRow()))
                        .collect(Collectors.toList());
                Range current = sorted.get(0);
                for (int i = 1; i < sorted.size(); i++) {
                    Range next = sorted.get(i);
                    if (current.getEndAddress().getRow() + 1 >= next.getStartAddress().getRow()) {
                        int newStartRow = current.getStartAddress().getRow();
                        int newEndRow = Math.max(current.getEndAddress().getRow(), next.getEndAddress().getRow());
                        current = new Range(current.getStartAddress().getColumn(), newStartRow, current.getEndAddress().getColumn(), newEndRow);
                    }
                    else {
                        mergedRanges.add(current);
                        current = next;
                    }
                }
                mergedRanges.add(current);
            }
        }
        else if (strategy == RangeMergeStrategy.MERGE_ROWS) {
            Map<String, List<Range>> groups = ranges.stream()
                    .collect(Collectors.groupingBy(r -> r.getStartAddress().getRow() + "," + r.getEndAddress().getRow()));
            for (List<Range> group : groups.values()) {
                List<Range> sorted = group.stream()
                        .sorted(Comparator.comparingInt(r -> r.getStartAddress().getColumn()))
                        .collect(Collectors.toList());
                Range current = sorted.get(0);
                for (int i = 1; i < sorted.size(); i++) {
                    Range next = sorted.get(i);
                    if (current.getEndAddress().getColumn() + 1 >= next.getStartAddress().getColumn()) {
                        int newStartCol = current.getStartAddress().getColumn();
                        int newEndCol = Math.max(current.getEndAddress().getColumn(), next.getEndAddress().getColumn());
                        current = new Range(newStartCol, current.getStartAddress().getRow(), newEndCol, current.getEndAddress().getRow());
                    }
                    else {
                        mergedRanges.add(current);
                        current = next;
                    }
                }
                mergedRanges.add(current);
            }
        }
        return mergedRanges;
    }

    /**
     * Method to generate an Excel internal password hash to protect workbooks or worksheets.<br>
     * <b>WARNING!</b> Do not use this method to encrypt real passwords. This is only a minor security feature.
     *
     * @param password Password as plain text
     * @return Encoded password
     */
    public static String generatePasswordHash(String password) {
        if (ParserUtils.isNullOrEmpty(password)) {
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
     * Method to create a {@link Duration} object from hours, minutes and seconds
     *
     * @param hours   Number of Hours within the day
     * @param minutes Number of minutes within the hour
     * @param seconds Number of seconds within the minute
     * @return Duration object
     * @throws FormatException thrown if one of the components is invalid
     */
    public static Duration createDuration(int hours, int minutes, int seconds) {
        return createDuration(0, hours, minutes, seconds);
    }

    /**
     * Method to create a {@link Duration} object from days, hours, minutes and seconds
     *
     * @param days    Total number of days
     * @param hours   Number of Hours within the day
     * @param minutes Number of minutes within the hour
     * @param seconds Number of seconds within the minute
     * @return Duration object
     * @throws FormatException thrown if one of the components is invalid, or if the number of days exceeds the year 9999
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
