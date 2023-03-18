/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * The import options define global rules to import worksheets. The options are
 * mainly to override particular cell types (e.g. interpretation of dates as
 * numbers)
 */
public class ImportOptions {

	/**
	 * Default format if Date values are cast to strings
	 */
	public static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

	/**
	 * Default format if time (Duration) values are cast to strings
	 */
	public static final String DEFAULT_TIME_FORMAT = "HH:mm:ss";

	/**
	 * Default locale instance (en-US) used for time parsing, if no custom locale is
	 * defined
	 */
	public static final Locale DEFAULT_LOCALE = Locale.US;

	/**
	 * Global conversion types to enforce during the import. All types other than
	 * {@link GlobalType#Default} will override defined {@link ColumnType}s
	 */
	public enum GlobalType {
		/**
		 * No global strategy. All numbers are tried to be cast to the most suitable
		 * types
		 */
		Default,
		/**
		 * All numbers are cast to doubles
		 */
		AllNumbersToDouble,
		/**
		 * All numbers are cast to BigDecimal
		 */
		AllNumbersToBigDecimal,
		/**
		 * All numbers are cast to integers. Floating point numbers will be rounded
		 * (commercial rounding) to the nearest integer
		 */
		AllNumbersToInt,
		/**
		 * Every cell is cast to a string. An empty cell will be cast to an empty string
		 */
		EverythingToString
	}

	/**
	 * Column types to enforce during the import
	 */
	public enum ColumnType {
		/**
		 * Cells are tried to be imported as numbers (automatic determination of numeric
		 * type)
		 */
		Numeric,
		/**
		 * Cells are tried to be imported as numbers (enforcing double)
		 */
		Double,
		/**
		 * Cells are tried to be imported as numbers (enforcing BigDecimal)
		 */
		BigDecimal,
		/**
		 * Cells are tried to be imported as dates (Date). If values are provided as
		 * Strings, the String defined with
		 * {@link ImportOptions#setDateFormat(java.lang.String)} will be used as parsing
		 * pattern
		 */
		Date,
		/**
		 * Cells are tried to be imported as times (LocalTime)
		 */
		Time,
		/**
		 * Cells are tried to be imported as booleans
		 */
		Bool,
		/**
		 * Cells are all imported as strings, using the ToString() method
		 */
		String
	}

	private boolean enforceDateTimesAsNumbers = false;
	private boolean enforceEmptyValuesAsString = false;
	private boolean enforcePhoneticCharacterImport = false;
	private final Map<Integer, ColumnType> enforcedColumnTypes = new HashMap<>();
	private int enforcingStartRowNumber = 0;
	private GlobalType globalEnforcingType = GlobalType.Default;
	private String dateFormat;
	private String timeFormat;
	private SimpleDateFormat dateFormatter;
	private DateTimeFormatter timeFormatter;
	private Locale temporalLocale;

	/**
	 * Gets whether date or time values in the workbook are interpreted as numbers
	 *
	 * @return If true, date or time values (default format number 14 or 21) will be
	 *         interpreted as numeric values globally. This option overrules
	 *         possible column options, defined by
	 *         {@link #addEnforcedColumn(int, ColumnType)}
	 */
	public boolean isEnforceDateTimesAsNumbers() {
		return enforceDateTimesAsNumbers;
	}

	/**
	 * gets the type enforcing rules during import for particular columns
	 *
	 * @return Map of column numers and enforced types
	 */
	public Map<Integer, ColumnType> getEnforcedColumnTypes() {
		return enforcedColumnTypes;
	}

	/**
	 * gets the row number (zero-based) where enforcing rules are started to be
	 * applied. This is, for instance, to prevent enforcing in a header row
	 *
	 * @return Row number
	 */
	public int getEnforcingStartRowNumber() {
		return enforcingStartRowNumber;
	}

	/**
	 * Sets whether date or time values in the workbook are interpreted as numbers
	 *
	 * @param enforceDateTimesAsNumbers
	 *            If true, date or time values (default format number 14 or 21) will
	 *            be interpreted as numeric values globally. This option overrules
	 *            possible column options, defined by
	 *            {@link #addEnforcedColumn(int, ColumnType)}
	 */
	public void setEnforceDateTimesAsNumbers(boolean enforceDateTimesAsNumbers) {
		this.enforceDateTimesAsNumbers = enforceDateTimesAsNumbers;
	}

	/**
	 * Sets the row number (zero-based) where enforcing rules are started to be
	 * applied. This is, for instance, to prevent enforcing types in a header row.
	 * Any enforcing rule is skipped until this row number is reached
	 *
	 * @param enforcingStartRowNumber
	 *            Row number
	 */
	public void setEnforcingStartRowNumber(int enforcingStartRowNumber) {
		this.enforcingStartRowNumber = enforcingStartRowNumber;
	}

	/**
	 * Adds a type enforcing rule to the passed column address
	 *
	 * @param columnAddress
	 *            Column address (A to XFD)
	 * @param type
	 *            Type to be enforced on the column
	 */
	public void addEnforcedColumn(String columnAddress, ColumnType type) {
		this.enforcedColumnTypes.put(Cell.resolveColumn(columnAddress), type);
	}

	/**
	 * Adds a type enforcing rule to the passed column number (zero-based)
	 *
	 * @param columnNumber
	 *            Column number (0-16383)
	 * @param type
	 *            Type to be enforced on the column
	 */
	public void addEnforcedColumn(int columnNumber, ColumnType type) {
		this.enforcedColumnTypes.put(columnNumber, type);
	}

	/**
	 * Gets whether phonetic characters (like ruby characters / Furigana / Zhuyin
	 * fuhao) in strings are added in brackets after the transcribed symbols. By
	 * default, phonetic characters are removed from strings.
	 *
	 * @return If true, phonetic characters will be appended, otherwise discarded
	 */
	public boolean isEnforcePhoneticCharacterImport() {
		return enforcePhoneticCharacterImport;
	}

	/**
	 * Sets whether phonetic characters (like ruby characters / Furigana / Zhuyin
	 * fuhao) in strings are added in brackets after the transcribed symbols. By
	 * default, phonetic characters are removed from strings.
	 *
	 * @param enforcePhoneticCharacterImport
	 *            If true, phonetic characters will be appended, otherwise discarded
	 * @apiNote This option is not applicable to specific rows or a start column
	 *          (applied globally)
	 */
	public void setEnforcePhoneticCharacterImport(boolean enforcePhoneticCharacterImport) {
		this.enforcePhoneticCharacterImport = enforcePhoneticCharacterImport;
	}

	/**
	 * Gets whether empty cells are of the type Empty or String
	 *
	 * @return If true, empty cells will be interpreted as type of string with an
	 *         empty value. If false, the type will be Empty and the value null
	 */
	public boolean isEnforceEmptyValuesAsString() {
		return enforceEmptyValuesAsString;
	}

	/**
	 * Sets whether empty cells are of the type Empty or String
	 *
	 * @param enforceEmptyValuesAsString
	 *            If true, empty cells will be interpreted as type of string with an
	 *            empty value. If false, the type will be Empty and the value null
	 */
	public void setEnforceEmptyValuesAsString(boolean enforceEmptyValuesAsString) {
		this.enforceEmptyValuesAsString = enforceEmptyValuesAsString;
	}

	/**
	 * Gets the global strategy to handle cell values. The default will not enforce
	 * any general casting, beside defined values of
	 * {@link ImportOptions#setEnforceDateTimesAsNumbers(boolean)},
	 * {@link ImportOptions#setEnforceEmptyValuesAsString(boolean)} and
	 * {@link ImportOptions#addEnforcedColumn(int, ColumnType)}
	 *
	 * @return Global cast strategy on import
	 */
	public GlobalType getGlobalEnforcingType() {
		return globalEnforcingType;
	}

	/**
	 * Sets the global strategy to handle cell values. The default will not enforce
	 * any casting, beside defined values of
	 * {@link ImportOptions#setEnforceDateTimesAsNumbers(boolean)},
	 * {@link ImportOptions#setEnforceEmptyValuesAsString(boolean)} and
	 * {@link ImportOptions#addEnforcedColumn(int, ColumnType)}
	 *
	 * @param globalEnforcingType
	 *            Global cast strategy on import
	 */
	public void setGlobalEnforcingType(GlobalType globalEnforcingType) {
		this.globalEnforcingType = globalEnforcingType;
	}

	/**
	 * Gets the format if Date values are cast to Strings
	 *
	 * @return String format pattern
	 */
	public String getDateFormat() {
		return dateFormat;
	}

	/**
	 * Sets the format if Date values are cast to Strings
	 *
	 * @param dateFormat
	 *            String format pattern
	 */
	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
		if (dateFormat == null) {
			this.dateFormatter = new SimpleDateFormat(DEFAULT_DATE_FORMAT);
		}
		else {
			this.dateFormatter = new SimpleDateFormat(dateFormat);
		}
	}

	/**
	 * Gets the Locale instance, used to parse Duration objects from strings. If
	 * null, parsing will be tried with 'best effort'.
	 *
	 * @return Locale instance used for parsing
	 */
	public Locale getTemporalLocale() {
		return temporalLocale;
	}

	/**
	 * Sets the Locale instance, used to parse Duration objects from strings. If
	 * null, parsing will be tried with 'best effort'.
	 *
	 * @param temporalLocale
	 *            Locale instance used for parsing
	 */
	public void setTemporalLocale(Locale temporalLocale) {
		this.temporalLocale = temporalLocale;
	}

	/**
	 * Gets the Date formatter, set by the string of
	 * {@link ImportOptions#setDateFormat(String)}
	 *
	 * @return SimpleDateFormat instance
	 */
	public SimpleDateFormat getDateFormatter() {
		return this.dateFormatter;
	}

	/**
	 * Gets the format if Duration (time) values are cast to Strings
	 *
	 * @return String format pattern
	 */
	public String getTimeFormat() {
		return timeFormat;
	}

	/**
	 * Sets the format if LocalTime values are cast to Strings.<br>
	 * Note that the parameter 'n' in a pattern is used to parse the number of days
	 * since 1900-01-01
	 *
	 * @param timeFormat
	 *            String format pattern
	 * @apiNote Supported formatting tokens are all time-related patterns like 'HH',
	 *          'mm', 'ss'. To represent the number of days, the pattern 'n' is
	 *          used. This deviates from the actual definition of 'n' which would be
	 *          nanoseconds of the second.
	 */
	public void setTimeFormat(String timeFormat) {
		this.timeFormat = timeFormat;
		if (timeFormat == null) {
			this.timeFormatter = DateTimeFormatter.ofPattern(DEFAULT_TIME_FORMAT).withLocale(this.temporalLocale);
		}
		else {
			this.timeFormatter = DateTimeFormatter.ofPattern(timeFormat).withLocale(this.temporalLocale);
		}
	}

	/**
	 * Gets the formatter to parse a Duration (time), set by the string of
	 * {@link ImportOptions#setTimeFormat(String)}
	 *
	 * @return DateTimeFormatter instance
	 */
	public DateTimeFormatter getTimeFormatter() {
		return this.timeFormatter;
	}

	public ImportOptions() {
		this.dateFormat = DEFAULT_DATE_FORMAT;
		this.timeFormat = DEFAULT_TIME_FORMAT;
		this.temporalLocale = DEFAULT_LOCALE;
		this.dateFormatter = new SimpleDateFormat(this.dateFormat);
		this.timeFormatter = DateTimeFormatter.ofPattern(this.timeFormat).withLocale(this.temporalLocale);
	}
}
