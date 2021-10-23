/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

/**
 * The import options define global rules to import worksheets. The options are mainly to override particular cell types (e.g. interpretation of dates as numbers)
 */
public class ImportOptions {

    /**
     * Default format if Date values are cast to strings
     */
    public static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    /**
     * Default format if LocalTime values are cast to strings
     */
    public static final String DEFAULT_LOCALTIME_FORMAT = "HH:mm:ss";

    /**
     *  Global conversion types to enforce during the import. All types other than {@link GlobalType#Default} will override defined {@link ColumnType}s
     */
    public enum GlobalType{
        /**
         * No global strategy. All numbers are tried to be cast to the most suitable types
         */
        Default,
        /**
         * ll numbers are cast to doubles
         */
        AllNumbersToDouble,
        /**
         * All numbers are cast to integers. Floating point numbers will be rounded (commercial rounding) to the nearest integer
         */
        AllNumbersToInt,
        /**
         * Every cell is cast to a string
         */
        EverythingToString
    }

    /**
     * Column types to enforce during the import
     */
    public enum ColumnType{
        /**
         * Cells are tried to be imported as numbers (automatic determination of numeric type)
         */
        Numeric,
        /**
         * Cells are tried to be imported as numbers (enforcing double)
         */
        Double,
        /**
         * Cells are tried to be imported as dates (Date)
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
    private final Map<Integer, ColumnType> enforcedColumnTypes = new HashMap<>();
    private  int enforcingStartRowNumber = 0;
    private GlobalType globalEnforcingType = GlobalType.Default;
    private String dateFormat;
    private String localTimeFormat;
    private SimpleDateFormat dateFormatter;
    private DateTimeFormatter localTimeFormatter;

    /**
     * Gets whether date or time values in the workbook are interpreted as numbers
     * @return  If true, date or time values (default format number 14 or 21) will be interpreted as numeric values globally. This option overrules possible column options, defined by {@link #addEnforcedColumn(int, ColumnType)}
     */
    public boolean isEnforceDateTimesAsNumbers() {
        return enforceDateTimesAsNumbers;
    }

    /**
     * gets the type enforcing rules during import for particular columns
     * @return Map of column numers and enforced types
     */
    public Map<Integer, ColumnType> getEnforcedColumnTypes() {
        return enforcedColumnTypes;
    }

    /**
     * gets the row number (zero-based) where enforcing rules are started to be applied. This is, for instance, to prevent enforcing in a header row
     * @return Row number
     */
    public int getEnforcingStartRowNumber() {
        return enforcingStartRowNumber;
    }

    /**
     *  Sets whether date or time values in the workbook are interpreted as numbers
     * @param enforceDateTimesAsNumbers If true, date or time values (default format number 14 or 21) will be interpreted as numeric values globally. This option overrules possible column options, defined by {@link #addEnforcedColumn(int, ColumnType)}
     */
    public void setEnforceDateTimesAsNumbers(boolean enforceDateTimesAsNumbers) {
        this.enforceDateTimesAsNumbers = enforceDateTimesAsNumbers;
    }

    /**
     * Sets the row number (zero-based) where enforcing rules are started to be applied. This is, for instance, to prevent enforcing types in a header row. Any enforcing rule is skipped until this row number is reached
     * @param enforcingStartRowNumber Row number
     */
    public void setEnforcingStartRowNumber(int enforcingStartRowNumber) {
        this.enforcingStartRowNumber = enforcingStartRowNumber;
    }

    /**
     * Adds a type enforcing rule to the passed column address
     * @param columnAddress Column address (A to XFD)
     * @param type Type to be enforced on the column
     */
    public void addEnforcedColumn(String columnAddress, ColumnType type)
    {
        this.enforcedColumnTypes.put(Cell.resolveColumn(columnAddress), type);
    }

    /**
     * Adds a type enforcing rule to the passed column number (zero-based)
     * @param columnNumber >Column number (0-16383)
     * @param type Type to be enforced on the column
     */
    public void addEnforcedColumn(int columnNumber, ColumnType type)
    {
        this.enforcedColumnTypes.put(columnNumber, type);
    }

    /**
     *  Gets whether empty cells are of the type Empty or String
     * @return If true, empty cells will be interpreted as type of string with an empty value. If false, the type will be Empty and the value null
     */
    public boolean isEnforceEmptyValuesAsString() {
        return enforceEmptyValuesAsString;
    }

    /**
     * Sets whether empty cells are of the type Empty or String
     * @param enforceEmptyValuesAsString If true, empty cells will be interpreted as type of string with an empty value. If false, the type will be Empty and the value null
     */
    public void setEnforceEmptyValuesAsString(boolean enforceEmptyValuesAsString) {
        this.enforceEmptyValuesAsString = enforceEmptyValuesAsString;
    }

    /**
     * Gets the global strategy to handle cell values. The default will not enforce any general casting, besides {@link ImportOptions#setEnforceDateTimesAsNumbers(boolean)}, {@link ImportOptions#setEnforceEmptyValuesAsString(boolean)} and {@link ImportOptions#addEnforcedColumn(int, ColumnType)}
     * @return Global cast strategy on import
     */
    public GlobalType getGlobalEnforcingType() {
        return globalEnforcingType;
    }

    /**
     * Sets the global strategy to handle cell values. The default will not enforce any casting, besides {@link ImportOptions#setEnforceDateTimesAsNumbers(boolean)}, {@link ImportOptions#setEnforceEmptyValuesAsString(boolean)} and {@link ImportOptions#addEnforcedColumn(int, ColumnType)}
     * @param globalEnforcingType Global cast strategy on import
     */
    public void setGlobalEnforcingType(GlobalType globalEnforcingType) {
        this.globalEnforcingType = globalEnforcingType;
    }

    /**
     * Gets the format if Date values are cast to Strings
     * @return String format pattern
     */
    public String getDateFormat() {
        return dateFormat;
    }

    /**
     * Sets the format if Date values are cast to Strings
     * @param dateFormat String format pattern
     */
    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
        this.dateFormatter = new SimpleDateFormat(dateFormat);
    }

    /**
     * Gets the Date formatter, set by the string of {@link ImportOptions#setDateFormat(String)}
     * @return SimpleDateFormat instance
     */
    public SimpleDateFormat getDateFormatter(){
        return this.dateFormatter;
    }

    /**
     * Gets the format if LocalTime values are cast to Strings
     * @return String format pattern
     */
    public String getLocalTimeFormat() {
        return localTimeFormat;
    }

    /**
     * Sets the format if LocalTime values are cast to Strings
     * @param localTimeFormat String format pattern
     */
    public void setLocalTimeFormat(String localTimeFormat) {
        this.localTimeFormat = localTimeFormat;
        this.localTimeFormatter = DateTimeFormatter.ofPattern(localTimeFormat);
    }

    /**
     * Gets the LocalTime formatter, set by the string of {@link ImportOptions#setLocalTimeFormat(String)}
     * @return DateTimeFormatter instance
     */
    public DateTimeFormatter getLocalTimeFormatter(){
        return this.localTimeFormatter;
    }

    public ImportOptions() {
        this.dateFormat = DEFAULT_DATE_FORMAT;
        this.localTimeFormat = DEFAULT_LOCALTIME_FORMAT;
        this.dateFormatter = new SimpleDateFormat(this.dateFormat);
        this.localTimeFormatter = DateTimeFormatter.ofPattern(this.localTimeFormat);
    }
}
