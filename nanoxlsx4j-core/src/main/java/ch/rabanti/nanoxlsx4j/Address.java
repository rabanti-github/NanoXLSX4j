/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.util.Objects;

// Note: Record shows this as class definition in Javadoc

/**
 * Record representing the cell address as column and row (zero based)
 *
 * @param column Column number (zero based)
 * @param row    Row number (zero based)
 * @param type   referencing type of the address
 */
public record Address(int column, int row, Cell.AddressType type) implements Comparable<Address> {

    // Note: Record shows this as full constructor in Javadoc; Add parameters here as well

    /**
     * Constructor with row, column and type as arguments. All referencing modifiers ($) are ignored and only the
     * defined referencing type considered
     *
     * @param column Column number (zero based)
     * @param row    Row number (zero based)
     * @param type   referencing type of the address
     */
    public Address {
        Cell.validateColumnNumber(column);
        Cell.validateRowNumber(row);
        Objects.requireNonNull(type, "type");
    }

    /**
     * Constructor with row and column as arguments.  The referencing type of the address is default (e.g. 'C20')
     *
     * @param column Column number (zero based)
     * @param row    Row number (zero based)
     */
    public Address(int column, int row) {
        this(column, row, Cell.AddressType.DEFAULT);
    }

    /**
     * Creates an address from an address string, retaining its referencing modifiers.
     *
     * @param address address string, for example {@code $B$12}
     */
    public Address(String address) {
        this(Cell.resolveCellCoordinate(address));
    }

    /**
     * Creates an address from an address string using the supplied referencing type. Referencing modifiers in the
     * string are ignored.
     *
     * @param address address string, for example {@code B12}
     * @param type    referencing type to use
     */
    public Address(String address, Cell.AddressType type) {
        this(Cell.resolveCellCoordinate(address), type);
    }

    /**
     * Internal constructor, using {@link Cell#resolveCellCoordinate(String)} directly
     *
     * @param address Address object
     */
    private Address(Address address) {
        this(address.column, address.row, address.type);
    }

    /**
     * Internal constructor, using main constructor directly
     *
     * @param address Address object
     * @param type    Override address type
     */
    private Address(Address address, Cell.AddressType type) {
        this(address.column, address.row, type);
    }

    /**
     * Gets the combined cell address.
     *
     * @return address in the format A1 - XFD1048576
     */
    public String getAddress() {
        return Cell.resolveCellAddress(column, row, type);
    }


   /**
    * Gets the column address (A - XFD)
    *
    * @return Column address as letter(s)
    */
   public String getColumn() {
       return Cell.resolveColumnAddress(column);
   }

   // /**
   //  * Gets the column number (zero based)
   //  * @return Column number as int
   //  */
   // public int getColumn() {
   //     return column;
   // }
//
   // /**
   //  * Gets the row number (zero based)
   //  * @return Row number
   //  */
   // public int getRow() {
   //     return row;
   // }

    /**
     * Gets the referencing type of the address
     * @return Referencing type
     */
    public Cell.AddressType getType() {
        return type;
    }

    /**
     * Overwritten ToString method
     *
     * @return Returns the cell address (e.g. 'A15')
     */
    @Override
    public String toString() {
        return getAddress();
    }

    /**
     * Gets the hash code based on the string representation of the address
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return toString().hashCode();
    }

    /**
     * Compares two objects whether they are addresses and equal
     *
     * @param o the reference object with which to compare.
     * @return True if equal and of the same type
     */
    @Override
    public boolean equals(Object o) {
        if (!(o instanceof Address(int column1, int row1, Cell.AddressType type1))) {
            return false;
        }

        return row == row1 && column == column1 && type == type1;
    }

    /**
     * Compares two Addresses whether they are equal
     *
     * @param other the object to be compared.
     * @return True if equal
     */
    @Override
    public int compareTo(Address other) {
        long coordinate = (long) column * Worksheet.MAX_ROW_NUMBER + row;
        long otherCoordinate = (long) other.column * Worksheet.MAX_ROW_NUMBER + other.row;
        return Long.compare(coordinate, otherCoordinate);
    }

    /**
     * Creates a (dereferenced, if applicable) deep copy of this address
     *
     * @return Copy of this address
     */
    Address copy() {
        return new Address(column, row, type);
    }
}
