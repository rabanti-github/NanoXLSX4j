//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.options;

import ch.rabanti.nanoxlsx4j.interfaces.IOptions;
import ch.rabanti.nanoxlsx4j.interfaces.ITextOptions;

import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * The reader options define global rules, applied when loading a worksheet.
 * The options are mainly to override particular cell types
 * (e.g. interpretation of dates as numbers).
 * <p>
 * Java equivalent of the .NET {@code ReaderOptions} class (formerly {@code ImportOptions}).
 * </p>
 */
public class ReaderOptions implements IOptions, ITextOptions {

    /**
     * Default format if DateTime values are cast to strings
     */
    public static final String DEFAULT_DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

    /**
     * Default format if Duration (TimeSpan) values are cast to strings
     */
    public static final String DEFAULT_TIME_SPAN_FORMAT = "HH:mm:ss";

    /**
     * Default locale instance (language-neutral / invariant culture) used for
     * date and time parsing, if no custom locale is defined.
     * <p>
     * Java equivalent of {@code CultureInfo.InvariantCulture}.
     * </p>
     */
    public static final Locale DEFAULT_CULTURE_INFO = Locale.ROOT;

    /**
     * Global conversion types to enforce during the load process.
     * All types other than {@link GlobalType#Default} will override defined {@link ColumnType}s.
     */
    public enum GlobalType {
        /**
         * No global strategy. All numbers are tried to be cast to the most suitable types
         */
        Default,
        /**
         * All numbers are cast to doubles
         */
        AllNumbersToDouble,
        /**
         * All numbers are cast to BigDecimal
         */
        AllNumbersToDecimal,
        /**
         * All numbers are cast to integers. Floating point numbers will be rounded
         * (commercial rounding) to the nearest integer
         */
        AllNumbersToInt,
        /**
         * Every cell is cast to a string. An empty cell will be cast to an empty string
         */
        EverythingToString,
    }

    /**
     * Column types to enforce during the read process.
     * The types are tried to be applied on all cells of a particular column.
     */
    public enum ColumnType {
        /**
         * Cells are tried to be imported as numbers (automatic determination of numeric type)
         */
        Numeric,
        /**
         * Cells are tried to be imported as numbers (enforcing double)
         */
        Double,
        /**
         * Cells are tried to be imported as numbers (enforcing BigDecimal)
         */
        Decimal,
        /**
         * Cells are tried to be imported as dates (Date). If values are provided as Strings,
         * the pattern defined with {@link ReaderOptions#setDateTimeFormat(String)} will be used
         */
        Date,
        /**
         * Cells are tried to be imported as times (Duration / LocalTime)
         */
        Time,
        /**
         * Cells are tried to be imported as booleans
         */
        Bool,
        /**
         * Cells are all imported as strings, using the toString() method
         */
        String
    }

    private boolean enforceDateTimesAsNumbers = false;
    private boolean enforceEmptyValuesAsString = false;
    private boolean enforcePhoneticCharacterImport = false;
    private boolean enforceStrictValidation = false;
    private boolean ignoreNotSupportedPasswordAlgorithms = false;
    private final Map<Integer, ColumnType> enforcedColumnTypes = new HashMap<>();
    private int enforcingStartRowNumber = 0;
    private GlobalType globalEnforcingType = GlobalType.Default;
    private String dateTimeFormat;
    private String timeSpanFormat;
    private SimpleDateFormat dateTimeFormatter;
    private DateTimeFormatter timeSpanFormatter;
    private Locale temporalCultureInfo;

    // -------------------------------------------------------------------------
    // Getters / setters — boolean flags
    // -------------------------------------------------------------------------

    /**
     * Gets whether date or time values in the workbook are interpreted as numbers.
     *
     * @return If true, date or time values (default format number 14 or 21) will be
     * interpreted as numeric values globally. This option overrules possible column
     * options defined by {@link #addEnforcedColumn(int, ColumnType)}
     */
    public boolean isEnforceDateTimesAsNumbers() {
        return enforceDateTimesAsNumbers;
    }

    /**
     * Sets whether date or time values in the workbook are interpreted as numbers.
     *
     * @param enforceDateTimesAsNumbers If true, date or time values (default format number
     *                                  14 or 21) will be interpreted as numeric values globally
     */
    public void setEnforceDateTimesAsNumbers(boolean enforceDateTimesAsNumbers) {
        this.enforceDateTimesAsNumbers = enforceDateTimesAsNumbers;
    }

    /**
     * Gets whether empty cells are of the type Empty or String.
     *
     * @return If true, empty cells will be interpreted as type string with an empty value.
     * If false, the type will be Empty and the value null
     */
    @Override
    public boolean isEnforceEmptyValuesAsString() {
        return enforceEmptyValuesAsString;
    }

    /**
     * Sets whether empty cells are of the type Empty or String.
     *
     * @param enforceEmptyValuesAsString If true, empty cells will be interpreted as type
     *                                   string with an empty value
     */
    @Override
    public void setEnforceEmptyValuesAsString(boolean enforceEmptyValuesAsString) {
        this.enforceEmptyValuesAsString = enforceEmptyValuesAsString;
    }

    /**
     * Gets whether phonetic characters (like ruby characters / Furigana / Zhuyin fuhao)
     * in strings are added in brackets after the transcribed symbols.
     * By default, phonetic characters are removed from strings.
     *
     * @return If true, phonetic characters will be appended, otherwise discarded
     * @apiNote This option is not applicable to specific rows or a start column (applied globally)
     */
    @Override
    public boolean isEnforcePhoneticCharacterImport() {
        return enforcePhoneticCharacterImport;
    }

    /**
     * Sets whether phonetic characters (like ruby characters / Furigana / Zhuyin fuhao)
     * in strings are added in brackets after the transcribed symbols.
     *
     * @param enforcePhoneticCharacterImport If true, phonetic characters will be appended,
     *                                       otherwise discarded
     * @apiNote This option is not applicable to specific rows or a start column (applied globally)
     */
    @Override
    public void setEnforcePhoneticCharacterImport(boolean enforcePhoneticCharacterImport) {
        this.enforcePhoneticCharacterImport = enforcePhoneticCharacterImport;
    }

    /**
     * Gets whether invalid data such as column widths or row heights that are out of range
     * will cause an exception when such a workbook is loaded.
     * When false (default), invalid values are silently clamped to the allowed minimum or maximum.
     *
     * @return If true, any out-of-range dimension will throw an exception
     */
    public boolean isEnforceStrictValidation() {
        return enforceStrictValidation;
    }

    /**
     * Sets whether invalid data such as column widths or row heights that are out of range
     * will cause an exception when such a workbook is loaded.
     *
     * @param enforceStrictValidation If true, any out-of-range dimension will throw an exception.
     *                                If false (default), values are silently clamped
     */
    public void setEnforceStrictValidation(boolean enforceStrictValidation) {
        this.enforceStrictValidation = enforceStrictValidation;
    }

    /**
     * Gets whether worksheet or workbook protection passwords with unknown / unsupported
     * algorithms will be ignored. If false (default), a
     * {@code NotSupportedContentException} will be thrown when such a hash is encountered.
     *
     * @return If true, unsupported password algorithm hashes are silently skipped
     */
    public boolean isIgnoreNotSupportedPasswordAlgorithms() {
        return ignoreNotSupportedPasswordAlgorithms;
    }

    /**
     * Sets whether worksheet or workbook protection passwords with unknown / unsupported
     * algorithms will be ignored.
     *
     * @param ignoreNotSupportedPasswordAlgorithms If true, unsupported password algorithm
     *                                             hashes are silently skipped. Default is false
     */
    public void setIgnoreNotSupportedPasswordAlgorithms(boolean ignoreNotSupportedPasswordAlgorithms) {
        this.ignoreNotSupportedPasswordAlgorithms = ignoreNotSupportedPasswordAlgorithms;
    }

    // -------------------------------------------------------------------------
    // Getters / setters — numeric / enum options
    // -------------------------------------------------------------------------

    /**
     * Gets the type enforcing rules during import for particular columns.
     *
     * @return Map of zero-based column numbers and enforced types
     */
    public Map<Integer, ColumnType> getEnforcedColumnTypes() {
        return enforcedColumnTypes;
    }

    /**
     * Gets the row number (zero-based) where enforcing rules start to be applied.
     * This is used, for instance, to prevent enforcing types in a header row.
     *
     * @return Zero-based row number
     */
    public int getEnforcingStartRowNumber() {
        return enforcingStartRowNumber;
    }

    /**
     * Sets the row number (zero-based) where enforcing rules start to be applied.
     * Any enforcing rule is skipped until this row number is reached.
     *
     * @param enforcingStartRowNumber Zero-based row number
     */
    public void setEnforcingStartRowNumber(int enforcingStartRowNumber) {
        this.enforcingStartRowNumber = enforcingStartRowNumber;
    }

    /**
     * Gets the global strategy to handle cell values on import.
     *
     * @return Global cast strategy on import
     */
    public GlobalType getGlobalEnforcingType() {
        return globalEnforcingType;
    }

    /**
     * Sets the global strategy to handle cell values on import.
     * The default will not enforce any general casting, beside values configured via
     * {@link #setEnforceDateTimesAsNumbers(boolean)}, {@link #setEnforceEmptyValuesAsString(boolean)}
     * and {@link #addEnforcedColumn(int, ColumnType)}.
     *
     * @param globalEnforcingType Global cast strategy on import
     */
    public void setGlobalEnforcingType(GlobalType globalEnforcingType) {
        this.globalEnforcingType = globalEnforcingType;
    }

    // -------------------------------------------------------------------------
    // Getters / setters — date / time format
    // -------------------------------------------------------------------------

    /**
     * Gets the format pattern used when DateTime values are cast to strings or
     * parsed from strings.
     *
     * @return String format pattern
     */
    public String getDateTimeFormat() {
        return dateTimeFormat;
    }

    /**
     * Sets the format pattern used when DateTime values are cast to strings or
     * parsed from strings. Passing {@code null} resets the formatter to the default pattern.
     *
     * @param dateTimeFormat String format pattern, or {@code null} to use the default
     */
    public void setDateTimeFormat(String dateTimeFormat) {
        this.dateTimeFormat = dateTimeFormat;
        if (dateTimeFormat == null) {
            this.dateTimeFormatter = new SimpleDateFormat(DEFAULT_DATE_TIME_FORMAT);
        }
        else {
            this.dateTimeFormatter = new SimpleDateFormat(dateTimeFormat);
        }
    }

    /**
     * Gets the format pattern used when Duration (TimeSpan) values are cast to strings.
     *
     * @return String format pattern
     */
    public String getTimeSpanFormat() {
        return timeSpanFormat;
    }

    /**
     * Sets the format pattern used when Duration (TimeSpan) values are cast to strings
     * or parsed from strings. Passing {@code null} resets the formatter to the default pattern.
     *
     * @param timeSpanFormat String format pattern, or {@code null} to use the default
     * @apiNote Supported tokens are all time-related patterns like {@code HH}, {@code mm},
     * {@code ss}. The token {@code n} represents the number of days (deviating from
     * the standard meaning of nanoseconds)
     */
    public void setTimeSpanFormat(String timeSpanFormat) {
        this.timeSpanFormat = timeSpanFormat;
        if (timeSpanFormat == null) {
            this.timeSpanFormatter = DateTimeFormatter.ofPattern(DEFAULT_TIME_SPAN_FORMAT).withLocale(this.temporalCultureInfo);
        }
        else {
            this.timeSpanFormatter = DateTimeFormatter.ofPattern(timeSpanFormat).withLocale(this.temporalCultureInfo);
        }
    }

    /**
     * Gets the locale instance used to parse Duration objects from strings.
     * If {@code null}, parsing will be tried with best-effort.
     * <p>
     * Java equivalent of .NET {@code TemporalCultureInfo}.
     * </p>
     *
     * @return Locale instance used for parsing
     */
    public Locale getTemporalCultureInfo() {
        return temporalCultureInfo;
    }

    /**
     * Sets the locale instance used to parse Duration objects from strings.
     *
     * @param temporalCultureInfo Locale instance, or {@code null} for best-effort parsing
     */
    public void setTemporalCultureInfo(Locale temporalCultureInfo) {
        this.temporalCultureInfo = temporalCultureInfo;
    }

    /**
     * Gets the {@link SimpleDateFormat} instance derived from the pattern set by
     * {@link #setDateTimeFormat(String)}.
     *
     * @return SimpleDateFormat instance
     */
    public SimpleDateFormat getDateTimeFormatter() {
        return this.dateTimeFormatter;
    }

    /**
     * Gets the {@link DateTimeFormatter} instance derived from the pattern set by
     * {@link #setTimeSpanFormat(String)}.
     *
     * @return DateTimeFormatter instance
     */
    public DateTimeFormatter getTimeSpanFormatter() {
        return this.timeSpanFormatter;
    }

    // -------------------------------------------------------------------------
    // Column-enforcing helpers
    // -------------------------------------------------------------------------

    /**
     * Adds a type enforcing rule for the column at the given address.
     *
     * @param columnAddress Column address (A to XFD)
     * @param type          Type to be enforced on the column
     */
    public void addEnforcedColumn(String columnAddress, ColumnType type) {
        this.enforcedColumnTypes.put(resolveColumnAddress(columnAddress), type);
    }

    /**
     * Adds a type enforcing rule for the column at the given zero-based index.
     *
     * @param columnNumber Zero-based column number (0–16383)
     * @param type         Type to be enforced on the column
     */
    public void addEnforcedColumn(int columnNumber, ColumnType type) {
        this.enforcedColumnTypes.put(columnNumber, type);
    }

    /**
     * Converts a column address string (e.g. {@code "A"}, {@code "AA"}) to a
     * zero-based column index.
     *
     * @param address Column address string (case-insensitive)
     * @return Zero-based column index
     */
    private static int resolveColumnAddress(String address) {
        String upper = address.toUpperCase(Locale.ROOT);
        int result = 0;
        for (char c : upper.toCharArray()) {
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1;
    }

    // -------------------------------------------------------------------------
    // Constructor
    // -------------------------------------------------------------------------

    /**
     * Default constructor — initialises all options to their defaults.
     */
    public ReaderOptions() {
        this.dateTimeFormat = DEFAULT_DATE_TIME_FORMAT;
        this.timeSpanFormat = DEFAULT_TIME_SPAN_FORMAT;
        this.temporalCultureInfo = DEFAULT_CULTURE_INFO;
        this.dateTimeFormatter = new SimpleDateFormat(this.dateTimeFormat);
        this.timeSpanFormatter = DateTimeFormatter.ofPattern(this.timeSpanFormat).withLocale(this.temporalCultureInfo);
    }
}
