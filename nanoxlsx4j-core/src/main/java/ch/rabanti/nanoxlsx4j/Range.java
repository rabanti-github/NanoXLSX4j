/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

// Note: Record shows this as class definition in Javadoc

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Record representing a cell range with a start and end address
 *
 * @param startAddress start address of the range
 * @param endAddress end address of the range
 */
public record Range(Address startAddress, Address endAddress) {

    // Note: Record shows this as full constructor in Javadoc; Add parameters here as well

    /**
     * Constructor with addresses as arguments. The addresses are automatically swapped if the start address is greater than the end address.
     * Referencing modifiers ($) for rows and columns can be passed through the address type of the address objects
     *  @param startAddress start address of the range
     *  @param endAddress end address of the range
     */
    public Range {
        if (startAddress.compareTo(endAddress) >= 0) {
            Address originalStart = startAddress;
            startAddress = endAddress;
            endAddress = originalStart;
        }
    }

    /**
     * Constructor with start and end rows and columns as arguments. The addresses are automatically swapped if the start address is greater than the end address.
     * Referencing modifiers ($) for rows and columns are not considered
     *
     * @param startColumn Start column number (zero based) of the range
     * @param startRow Start row number (zero based) of the range
     * @param endColumn End column number (zero based) of the range
     * @param endRow End row number (zero based) of the range
     */
    public Range(int startColumn, int startRow, int endColumn, int endRow){
        this(new  Address(startColumn, startRow), new Address(endColumn, endRow));
    }

    /**
     * Constructor with a range string as argument. The addresses are automatically swapped if the start address is greater than the end address.
     * Referencing modifiers ($) for rows and columns can be defined in the passed string
     *
    * @param range Address range (e.g. 'A1:B12')
     */
    public Range(String range) {
        this(Cell.resolveCellRange(range));
    }

    private Range(Range range) {
        this(range.startAddress, range.endAddress);
    }

    /**
     * Gets whether another range is completely enclosed by this range
     *
     * @param other Other range to check
     * @return True if the other range is completely enclosed. False if only partial overlapping or not intersecting
     */
    public boolean contains(Range other)
    {
        return this.startAddress.column() <= other.startAddress.column() &&
                this.endAddress.column() >= other.endAddress.column() &&
                this.startAddress.row() <= other.startAddress.row() &&
                this.endAddress.row() >= other.endAddress.row();
    }

    /**
     * Determines whether an address is within this range
     *
     * @param address Address to check
     * @return True if the address is part of this range, otherwise false
     */
    public boolean contains(Address address)
    {
        return address.column() >= this.startAddress.column() &&
                address.column() <= this.endAddress.column() &&
                address.row() >= this.startAddress.row() &&
                address.row() <= this.endAddress.row();
    }

    /**
     * Determines whether the passed range overlaps with this range
     *
     * @param other Range to check for overlapping
     * @return True if overlapping, otherwise false
     */
    public boolean Overlaps(Range other)
    {
        return !(this.endAddress.row() < other.startAddress.row() || this.startAddress.row() > other.endAddress.row() ||
                this.endAddress.column() < other.startAddress.column() || this.startAddress.column() > other.endAddress.column());
    }

    /**
     * Gets a list of all addresses between the start and end address
     *
     * <p>Remarks: Use this function with caution. Very big ranges may result to hundred of Millions or even Billions of cells. This may lead to an extremely high memory consumptions or even a crash of the application</p>
     *
     * @return List of Addresses
     */
    public List<Address> resolveEnclosedAddresses()
    {
        List<Address> range = Cell.getCellRange(this.startAddress, this.endAddress);
        return new ArrayList<>(range);
    }

    /**
     * Overwritten ToString method
     * @return Returns the range (e.g. 'A1:B12')
     */
    @Override
    public String toString() {
        return startAddress + ":" + endAddress;
    }

    /**
     * Compares two objects whether they are ranges and equal. The cell types (possible $ prefix) are considered
     * @param o   the reference object with which to compare.
     * @return True if the two objects are the same range
     */
    @Override
    public boolean equals(Object o) {
        if (!(o instanceof Range range))
            return false;

        return Objects.equals(endAddress, range.endAddress) && Objects.equals(startAddress, range.startAddress);
    }

    /**
     * Gets the hash code of the range object according to its string representation
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return this.toString().hashCode();
    }
}
