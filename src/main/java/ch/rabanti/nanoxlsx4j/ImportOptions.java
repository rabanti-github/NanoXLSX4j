/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.util.HashMap;
import java.util.Map;

/**
 * The import options define global rules to import worksheets. The options are mainly to override particular cell types (e.g. interpretation of dates as numbers)
 */
public class ImportOptions {

    /**
     * Column types to enforce during the import
     */
    public enum ColumnType{
        /**
         * Cells are tried to be imported as numbers (double)
         */
        numeric,
        /**
         * Cells are tried to be imported as dates (Date)
         */
        date,
        /**
         * Cells are tried to be imported as times (LocalTime)
         */
        time,
        /**
         * Cells are tried to be imported as booleans
         */
        bool,
        /**
         * Cells are all imported as strings, using the ToString() method
         */
        string
    }

    private boolean enforceDateTimesAsNumbers = false;
    private Map<Integer, ColumnType> enforcedColumnTypes = new HashMap<>();
    private  int EnforcingStartRowNumber = 0;

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
        return EnforcingStartRowNumber;
    }

    /**
     *  Sets whether date or time values in the workbook are interpreted as numbers
     * @param enforceDateTimesAsNumbers If true, date or time values (default format number 14 or 21) will be interpreted as numeric values globally. This option overrules possible column options, defined by {@link #addEnforcedColumn(int, ColumnType)}
     */
    public void setEnforceDateTimesAsNumbers(boolean enforceDateTimesAsNumbers) {
        this.enforceDateTimesAsNumbers = enforceDateTimesAsNumbers;
    }

    /**
     * Sets the row number (zero-based) where enforcing rules are started to be applied. This is, for instance, to prevent enforcing in a header row
     * @param enforcingStartRowNumber Row number
     */
    public void setEnforcingStartRowNumber(int enforcingStartRowNumber) {
        EnforcingStartRowNumber = enforcingStartRowNumber;
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


}
