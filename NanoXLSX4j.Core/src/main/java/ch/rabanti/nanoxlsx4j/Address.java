/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

import static ch.rabanti.nanoxlsx4j.Cell.resolveCellAddress;

/**
 * Class representing a cell address as column and row (zero-based)
 *
 * @author Raphael Stoeckli
 */
public class Address implements Comparable<Address> {

    // ### P R I V A T E  F I E L D S ###

    private final int column;
    private final int row;
    private final Cell.AddressType type;

    // ### G E T T E R S ###

    /**
     * Gets the column number (zero-based)
     *
     * @return Column number
     */
    public int getColumn() {
        return column;
    }

    /**
     * Gets the row number (zero-based)
     *
     * @return Row number
     */
    public int getRow() {
        return row;
    }

    /**
     * Gets the referencing type of the address
     *
     * @return Address referencing type
     */
    public Cell.AddressType getType() {
        return type;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Constructor with column and row as arguments. The referencing type of the address is default (e.g. 'C20')
     *
     * @param column Column number (zero-based)
     * @param row    Row number (zero-based)
     * @throws RangeException Thrown if the resolved address is out of range
     */
    public Address(int column, int row) {
        this(column, row, Cell.AddressType.Default);
    }

    /**
     * Constructor with column, row and address type as arguments
     *
     * @param column Column number (zero-based)
     * @param row    Row number (zero-based)
     * @param type   Referencing type of the address
     * @throws RangeException Thrown if the resolved address is out of range
     */
    public Address(int column, int row, Cell.AddressType type) {
        Cell.validateColumnNumber(column);
        Cell.validateRowNumber(row);
        this.column = column;
        this.row = row;
        this.type = type;
    }

    /**
     * Constructor with address as string. If no referencing modifiers ($) are defined, the address is of referencing
     * type default (e.g. 'C23')
     *
     * @param address Address string (e.g. '$B$12')
     * @throws RangeException Thrown if the resolved address is out of range
     */
    public Address(String address) {
        Address a = Cell.resolveCellCoordinate(address);
        this.column = a.column;
        this.row = a.row;
        this.type = a.type;
    }

    /**
     * Constructor with address as string. All referencing modifiers ($) are ignored and only the defined referencing
     * type is considered
     *
     * @param address Address string (e.g. 'B12')
     * @param type    Referencing type of the address
     * @throws RangeException Thrown if the resolved address is out of range
     */
    public Address(String address, Cell.AddressType type) {
        this.type = type;
        Address a = Cell.resolveCellCoordinate(address);
        this.column = a.column;
        this.row = a.row;
    }

    // ### M E T H O D S ###

    /**
     * Returns the combined address as string (e.g. 'A1' - 'XFD1048576')
     *
     * @return Address as string
     * @throws RangeException Thrown if the column or row is out of range
     */
    public String getAddress() {
        return resolveCellAddress(this.column, this.row, this.type);
    }

    /**
     * Gets the column address as letters (A - XFD)
     *
     * @return Column address as letter(s)
     */
    public String getColumnAddress() {
        return Cell.resolveColumnAddress(column);
    }

    /**
     * Returns the cell address as string (e.g. 'A15')
     *
     * @return Address as string
     */
    @Override
    public String toString() {
        return getAddress();
    }

    /**
     * Compares two addresses whether they are equal (row, column and referencing type must match)
     *
     * @param o Address to compare
     * @return True if equal
     */
    @Override
    public boolean equals(Object o) {
        if (o == this) {
            return true;
        }
        if (!(o instanceof Address)) {
            return false;
        }
        Address address = (Address) o;
        return this.row == address.row && this.column == address.column && this.type == address.type;
    }

    /**
     * Gets the hash code of the address based on its string representation
     *
     * @return Hash code of the address
     */
    @Override
    public int hashCode() {
        return this.toString().hashCode();
    }

    /**
     * Compares two addresses using the column and row numbers
     *
     * @param other Other address to compare
     * @return Negative if this address is smaller, 0 if equal, positive if this address is greater
     */
    @Override
    public int compareTo(Address other) {
        long thisCoordinate = (long) column * (long) Worksheet.MAX_ROW_NUMBER + row;
        long otherCoordinate = (long) other.column * (long) Worksheet.MAX_ROW_NUMBER + other.row;
        return Long.compare(thisCoordinate, otherCoordinate);
    }

    /**
     * Creates a (dereferenced, if applicable) deep copy of this address
     *
     * @return Copy of this address
     */
    Address copy() {
        return new Address(this.column, this.row, this.type);
    }

}
