/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

/**
 * Class representing a column of a worksheet
 */
public class Column {

    private int number;
    private String columnAddress;
    private float width;
    private Style defaultColumnStyle;
    private boolean autoFilter;
    private boolean hidden;

    /**
     * Gets the column address (A to XFD)
     *
     * @return Column address as string
     */
    public String getColumnAddress() {
        return columnAddress;
    }

    public void setColumnAddress(String columnAddress) {
        if (ParserUtils.isNullOrEmpty(columnAddress)) {
            throw new RangeException("The passed address was null or empty");
        }
        number = Cell.resolveColumn(columnAddress);
        this.columnAddress = ParserUtils.toUpper(columnAddress);
    }

    /**
     * Gets whether the column has auto filter applied
     *
     * @return If true, the column has auto filter applied, otherwise not
     */
    public boolean hasAutoFilter() {
        return autoFilter;
    }

    /**
     * Sets auto filter on the column or removes it
     *
     * @param autoFilter If true, the column has auto filter applied, otherwise not
     */
    public void setAutoFilter(boolean autoFilter) {
        this.autoFilter = autoFilter;
    }

    /**
     * Gets whether te column is hidden
     *
     * @return If true, the column is hidden, otherwise visible
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets the column to hidden or visible
     *
     * @param hidden If true, the column is hidden, otherwise visible
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    /**
     * Gets the column number (0 to 16383)
     *
     * @return Column number as int
     */
    public int getNumber() {
        return number;
    }

    /**
     * Sets the column number (0 to 16383)
     *
     * @param number Column number as int
     */
    public void setNumber(int number) {
        columnAddress = Cell.resolveColumnAddress(number);
        this.number = number;
    }

    /**
     * Gets the width of the column
     *
     * @return Width of the column
     */
    public float getWidth() {
        return width;
    }

    /**
     * Sets the width of the column
     *
     * @param width Width of the column
     */
    public void setWidth(float width) {
        if (width < Worksheet.MIN_COLUMN_WIDTH || width > Worksheet.MAX_COLUMN_WIDTH) {
            throw new RangeException("The passed column width is out of range (" + Worksheet.MIN_COLUMN_WIDTH + " to " +
                    Worksheet.MAX_COLUMN_WIDTH + ")");
        }
        this.width = width;
    }

    /**
     * Gets the default style of the column
     *
     * @return Column style
     */
    public Style getDefaultColumnStyle() {
        return defaultColumnStyle;
    }

    /**
     * Sets the default style of the column. This method may provide an updated style object as return value
     *
     * @param defaultColumnStyle Style to assign as default column style. Can be null (to clear)
     * @return If the passed style already exists in the repository, the existing one will be returned, otherwise the
     * passed one
     */

    public Style setDefaultColumnStyle(Style defaultColumnStyle) {
        return setDefaultColumnStyleInternal(defaultColumnStyle, false);
    }

    /**
     * Sets the default style of the column internally
     *
     * @param defaultColumnStyle Style to assign as default column style. Can be null (to clear)
     * @param unmanaged          Internal use only: If true, the style repository is not invoked and only the style
     *                           object of the cell is updated. Do not use!
     * @return If the passed style already exists in the repository, the existing one will be returned, otherwise the
     * passed one
     */
    Style setDefaultColumnStyleInternal(Style defaultColumnStyle, boolean unmanaged) {
        if (defaultColumnStyle == null) {
            this.defaultColumnStyle = null;
            return null;
        }
        if (unmanaged) {
            this.defaultColumnStyle = defaultColumnStyle;
        } else {
            this.defaultColumnStyle = StyleRepository.getInstance().addStyle(defaultColumnStyle);
        }
        return this.defaultColumnStyle;
    }

    /**
     * Default constructor (private, since not valid without address)
     */
    private Column() {
        this.width = Worksheet.DEFAULT_WORKSHEET_COLUMN_WIDTH;
        defaultColumnStyle = null;
    }

    /**
     * Constructor with column number
     *
     * @param columnCoordinate Column number (zero-based, 0 to 16383)
     */
    public Column(int columnCoordinate) {
        this();
        this.number = columnCoordinate;
    }

    /**
     * Constructor with column address
     *
     * @param columnAddress Column address (A to XFD)
     */
    public Column(String columnAddress) {
        this();
        this.columnAddress = columnAddress;
    }

    /**
     * Creates a deep copy of this column
     *
     * @return Copy of this column
     */
    Column copy() {
        Column copy = new Column();
        copy.hidden = this.hidden;
        copy.width = this.width;
        copy.autoFilter = this.autoFilter;
        copy.columnAddress = this.columnAddress;
        copy.number = this.number;
        copy.defaultColumnStyle = this.defaultColumnStyle;
        return copy;
    }

}
