/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.options;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.exceptions.NotSupportedContentException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Options;
import ch.rabanti.nanoxlsx4j.internal.interfaces.TextOptions;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

/**
 * Defines global rules applied while loading a worksheet. These options mainly override the interpretation of cell
 * values, for example by reading dates as numbers. Formula expressions and cached formula values are not converted.
 */
public class ReaderOptions implements Options, TextOptions {

    /** Default format used when date-time values are converted to strings. */
    public static final String DEFAULT_DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

    /** Default format used when duration values are converted to strings. */
    public static final String DEFAULT_TIME_SPAN_FORMAT = "hh\\:mm\\:ss";

    /** Default locale used for date and time parsing when no custom locale is defined. */
    public static final Locale DEFAULT_TEMPORAL_LOCALE = Locale.ROOT;

    /**
     * Global conversion types to enforce during the load process. Values other than {@link GlobalType#DEFAULT}
     * override rules defined by {@link ColumnType column types}. Formula cells are excluded from global conversion.
     */
    public enum GlobalType {
        /** No global strategy. Numbers are converted to the most suitable numeric types. */
        DEFAULT,
        /** All numbers are converted to {@link Double}. */
        ALL_NUMBERS_TO_DOUBLE,
        /** All numbers are converted to {@link java.math.BigDecimal}. */
        ALL_NUMBERS_TO_BIG_DECIMAL,
        /** Floating-point numbers are commercially rounded to the nearest {@link Integer}. */
        ALL_NUMBERS_TO_INT,
        /** Every non-formula cell is converted to a string. */
        EVERYTHING_TO_STRING
    }

    /** Conversion types applied to all non-formula cells of a particular column. */
    public enum ColumnType {
        /** Values are imported as numbers with automatic numeric-type selection. */
        NUMERIC,
        /** Values are imported as {@link Double} numbers. */
        DOUBLE,
        /** Values are imported as {@link java.math.BigDecimal} numbers. */
        BIG_DECIMAL,
        /**
         * Values are imported as dates. See {@link ReaderOptions#getDateTimeFormat()},
         * {@link ReaderOptions#getTimeSpanFormat()}, and {@link ReaderOptions#getTemporalLocale()}.
         */
        DATE,
        /** Values are imported as times or durations. */
        TIME,
        /** Values are imported as booleans. */
        BOOLEAN,
        /** Values are imported as strings using their string representation. */
        STRING
    }

    private boolean enforceDateTimesAsNumbers;
    private boolean enforcePhoneticCharacterImport;
    private boolean enforceEmptyValuesAsString;
    private boolean enforceStrictValidation;
    private GlobalType globalEnforcingType = GlobalType.DEFAULT;
    private final Map<Integer, ColumnType> enforcedColumnTypes = new HashMap<>();
    private int enforcingStartRowNumber;
    private String dateTimeFormat = DEFAULT_DATE_TIME_FORMAT;
    private String timeSpanFormat = DEFAULT_TIME_SPAN_FORMAT;
    private Locale temporalLocale = DEFAULT_TEMPORAL_LOCALE;
    private boolean ignoreNotSupportedPasswordAlgorithms;

    /** Creates reader options initialized with the NanoXLSX defaults. */
    public ReaderOptions() {
    }

    /**
     * Gets whether date or time values using the default number formats are interpreted as numbers globally. This
     * option overrides column rules defined by {@link ReaderOptions#addEnforcedColumn(int, ColumnType)}.
     *
     * @return true if date and time values are interpreted as numbers
     */
    public boolean getEnforceDateTimesAsNumbers() {
        return enforceDateTimesAsNumbers;
    }

    /**
     * Sets whether date and time values are interpreted as numbers.
     *
     * @param value true to interpret date and time values as numbers
     */
    public void setEnforceDateTimesAsNumbers(boolean value) {
        enforceDateTimesAsNumbers = value;
    }

    /** @return true if phonetic characters are appended after the transcribed symbols */
    @Override
    public boolean getEnforcePhoneticCharacterImport() {
        return enforcePhoneticCharacterImport;
    }

    /** @param value true to preserve phonetic characters during import; this option is applied globally */
    @Override
    public void setEnforcePhoneticCharacterImport(boolean value) {
        enforcePhoneticCharacterImport = value;
    }

    /** @return true if empty values are imported as strings instead of empty cells with null values */
    @Override
    public boolean getEnforceEmptyValuesAsString() {
        return enforceEmptyValuesAsString;
    }

    /** @param value true to import empty values as strings */
    @Override
    public void setEnforceEmptyValuesAsString(boolean value) {
        enforceEmptyValuesAsString = value;
    }

    /**
     * Gets whether strict validation is enabled.
     *
     * @return true if invalid workbook data causes an exception instead of being read in tolerant mode
     */
    public boolean getEnforceStrictValidation() {
        return enforceStrictValidation;
    }

    /**
     * Sets whether strict validation is enabled.
     *
     * @param value true to enable strict validation
     */
    public void setEnforceStrictValidation(boolean value) {
        enforceStrictValidation = value;
    }

    /**
     * Gets the global cell-value conversion strategy.
     *
     * @return global cell-value conversion strategy; the default is {@link GlobalType#DEFAULT}
     */
    public GlobalType getGlobalEnforcingType() {
        return globalEnforcingType;
    }

    /**
     * Sets the global cell-value conversion strategy.
     *
     * @param globalEnforcingType global cell-value conversion strategy
     */
    public void setGlobalEnforcingType(GlobalType globalEnforcingType) {
        this.globalEnforcingType = Objects.requireNonNull(globalEnforcingType, "globalEnforcingType");
    }

    /**
     * Gets the mutable mapping of zero-based column numbers to enforced conversion types. Formula cells are excluded
     * from column type enforcement.
     *
     * @return enforced column conversion rules
     */
    public Map<Integer, ColumnType> getEnforcedColumnTypes() {
        return enforcedColumnTypes;
    }

    /**
     * Gets the first row to which conversion rules apply.
     *
     * @return zero-based row number from which conversion rules are applied
     */
    public int getEnforcingStartRowNumber() {
        return enforcingStartRowNumber;
    }

    /**
     * Sets the first row to which conversion rules apply.
     *
     * @param enforcingStartRowNumber zero-based row number from which conversion rules are applied
     */
    public void setEnforcingStartRowNumber(int enforcingStartRowNumber) {
        this.enforcingStartRowNumber = enforcingStartRowNumber;
    }

    /**
     * Gets the format used when date-time values are converted to strings or parsed from strings. A null or empty value
     * requests best-effort parsing using the configured {@link ReaderOptions#getTemporalLocale()}.
     *
     * @return date-time format
     */
    public String getDateTimeFormat() {
        return dateTimeFormat;
    }

    /**
     * Sets the date-time format.
     *
     * @param dateTimeFormat date-time format, or null/empty for best-effort parsing
     */
    public void setDateTimeFormat(String dateTimeFormat) {
        this.dateTimeFormat = dateTimeFormat;
    }

    /**
     * Gets the duration format.
     *
     * @return format used when duration values are converted to strings
     */
    public String getTimeSpanFormat() {
        return timeSpanFormat;
    }

    /**
     * Sets the duration format.
     *
     * @param timeSpanFormat format used when duration values are converted to strings
     */
    public void setTimeSpanFormat(String timeSpanFormat) {
        this.timeSpanFormat = timeSpanFormat;
    }

    /**
     * Gets the locale used to parse date-time or duration values from strings. A null value requests best-effort
     * parsing.
     *
     * @return temporal parsing locale
     */
    public Locale getTemporalLocale() {
        return temporalLocale;
    }

    /**
     * Sets the temporal parsing locale.
     *
     * @param temporalLocale temporal parsing locale, or null for best-effort parsing
     */
    public void setTemporalLocale(Locale temporalLocale) {
        this.temporalLocale = temporalLocale;
    }

    /**
     * Gets whether protection passwords using unknown algorithms are ignored. Otherwise,
     * {@link NotSupportedContentException} is raised by the reader.
     *
     * @return true if unsupported password algorithms are ignored
     */
    public boolean getIgnoreNotSupportedPasswordAlgorithms() {
        return ignoreNotSupportedPasswordAlgorithms;
    }

    /**
     * Sets whether unsupported password algorithms are ignored.
     *
     * @param value true to ignore unsupported password algorithms
     */
    public void setIgnoreNotSupportedPasswordAlgorithms(boolean value) {
        ignoreNotSupportedPasswordAlgorithms = value;
    }

    /**
     * Adds a conversion rule for a column address.
     *
     * @param columnAddress column address from A to XFD
     * @param type conversion type enforced on the column
     * @throws ch.rabanti.nanoxlsx4j.exceptions.RangeException if the column address is outside the supported range
     * @throws IllegalArgumentException if a rule already exists for the resolved column
     */
    public void addEnforcedColumn(String columnAddress, ColumnType type) {
        addEnforcedColumn(Cell.resolveColumn(columnAddress), type);
    }

    /**
     * Adds a conversion rule for a zero-based column number.
     *
     * @param columnNumber zero-based column number
     * @param type conversion type enforced on the column
     * @throws IllegalArgumentException if a rule already exists for the column
     */
    public void addEnforcedColumn(int columnNumber, ColumnType type) {
        Objects.requireNonNull(type, "type");
        if (enforcedColumnTypes.containsKey(columnNumber)) {
            throw new IllegalArgumentException("A conversion rule already exists for column " + columnNumber);
        }
        enforcedColumnTypes.put(columnNumber, type);
    }
}
