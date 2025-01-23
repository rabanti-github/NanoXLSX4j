/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;

/**
 * Class representing a column of a worksheet
 *
 * @author Raphael Stoeckli
 */
public class Column {

    // ### P R I V A T E F I E L D S ###
    private boolean autoFilter;
    private String columnAddress;
    private boolean hidden;
    private int number;
    private float width;
    private Style defaultColumnStyle;

    // ### G E T T E R S & S E T T E R S ###

    /**
     * Sets whether auto filter is enabled on the column
     *
     * @param hasAutoFilter If true, the column has auto filter applied, otherwise not
     */
    public void setAutoFilter(boolean hasAutoFilter) {
        this.autoFilter = hasAutoFilter;
    }

    /**
     * Gets the column address
     *
     * @return Column address (A to XFD)
     */
    public String getColumnAddress() {
        return columnAddress;
    }

    /**
     * Sets the column address
     *
     * @param columnAddress Column address (A to XFD)
     */
    public void setColumnAddress(String columnAddress) {
        if (Helper.isNullOrEmpty(columnAddress)) {
            throw new RangeException("The passed address was null or empty");
        }
        this.number = Cell.resolveColumn(columnAddress);
        this.columnAddress = columnAddress.toUpperCase();
    }

    /**
     * Gets the column number
     *
     * @return Column number (0 to 16383)
     */
    public int getNumber() {
        return number;
    }

    /**
     * Sets the column number
     *
     * @param number Column number (0 to 16383)
     */
    public void setNumber(int number) {
        this.columnAddress = Cell.resolveColumnAddress(number);
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
            throw new RangeException("The passed column width is out of range (" + Worksheet.MIN_COLUMN_WIDTH + " to " + Worksheet.MAX_COLUMN_WIDTH + ")");
        }
        this.width = width;
    }

    /**
     * Gets whether auto filter is enabled on the column
     *
     * @return If true, the column has auto filter applied, otherwise not
     */
    public boolean hasAutoFilter() {
        return autoFilter;
    }

    /**
     * Gets whether the column is hidden or visible
     *
     * @return If true, the column is hidden, otherwise visible
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the column is hidden or visible
     *
     * @param isHidden If true, the column is hidden, otherwise visible
     */
    public void setHidden(boolean isHidden) {
        this.hidden = isHidden;
    }

    /**
     * Gets the default style of the column
     *
     * @return Default style
     */
    public Style getDefaultColumnStyle() {
        return defaultColumnStyle;
    }

    /**
     * Sets the default style of the column
     *
     * @param defaultColumnStyle Style to assign as default column style. Can be null (to clear)
     * @return If the passed style already exists in the repository, the existing one will be returned, otherwise the
     * passed one
     */
    public Style setDefaultColumnStyle(Style defaultColumnStyle) {
        return setDefaultColumnStyle(defaultColumnStyle, false);
    }

    /**
     * Sets the default style of the column
     *
     * @param defaultColumnStyle Style to assign as default column style. Can be null (to clear)
     * @param unmanaged          Internally used: If true, the style repository is not invoked and only the style object
     *                           of the column is updated. Do not use!
     * @return If the passed style already exists in the repository, the existing one will be returned, otherwise the
     * passed one
     */
    public Style setDefaultColumnStyle(Style defaultColumnStyle, boolean unmanaged) {
        if (defaultColumnStyle == null) {
            this.defaultColumnStyle = null;
            return null;
        }
        if (unmanaged) {
            this.defaultColumnStyle = defaultColumnStyle;
        }
        else {
            this.defaultColumnStyle = StyleRepository.getInstance().addStyle(defaultColumnStyle);
        }
        return this.defaultColumnStyle;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor (private, since not valid without address)
     */
    private Column() {
        this.width = Worksheet.DEFAULT_COLUMN_WIDTH;
        this.defaultColumnStyle = null;
    }

    /**
     * Constructor with column address
     *
     * @param columnAddress Column address (A to XFD)
     */
    public Column(String columnAddress) {
        this();
        this.setColumnAddress(columnAddress);
    }

    /**
     * Constructor with column number
     *
     * @param number Column number (zero-based, 0 to 16383)
     */
    public Column(int number) {
        this();
        this.setNumber(number);
    }

    /**
     * Creates a deep copy of this column
     *
     * @return Copy of this column
     */
    Column copy() {
        Column copy = new Column();
        copy.setHidden(this.hidden);
        copy.setWidth(this.width);
        copy.setAutoFilter(this.autoFilter);
        copy.columnAddress = this.columnAddress;
        copy.number = this.number;
        if (this.defaultColumnStyle != null) {
            copy.setDefaultColumnStyle(this.defaultColumnStyle, true);
        }
        return copy;
    }

}
