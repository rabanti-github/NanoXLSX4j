/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.util.List;

/**
 * Class representing a cell range with a start and end address
 *
 * @author Raphael Stoeckli
 */
public class Range {

    // ### P R I V A T E  F I E L D S ###

    private final Address startAddress;
    private final Address endAddress;

    // ### G E T T E R S ###

    /**
     * Gets the end address of the range
     *
     * @return End address
     */
    public Address getEndAddress() {
        return endAddress;
    }

    /**
     * Gets the start address of the range
     *
     * @return Start address
     */
    public Address getStartAddress() {
        return startAddress;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Constructor with addresses as arguments. The addresses are automatically swapped if the start address is
     * greater than the end address
     *
     * @param start Start address of the range
     * @param end   End address of the range
     */
    public Range(Address start, Address end) {
        if (start.compareTo(end) < 0) {
            this.startAddress = start;
            this.endAddress = end;
        }
        else {
            this.startAddress = end;
            this.endAddress = start;
        }
    }

    /**
     * Constructor with start and end rows and columns as arguments. The addresses are automatically swapped if
     * the start address is greater than the end address
     *
     * @param startColumn Start column number (zero-based) of the range
     * @param startRow    Start row number (zero-based) of the range
     * @param endColumn   End column number (zero-based) of the range
     * @param endRow      End row number (zero-based) of the range
     */
    public Range(int startColumn, int startRow, int endColumn, int endRow) {
        this(new Address(startColumn, startRow), new Address(endColumn, endRow));
    }

    /**
     * Constructor with a range string as argument. The addresses are automatically swapped if the start address is
     * greater than the end address
     *
     * @param range Address range (e.g. 'A1:B12')
     */
    public Range(String range) {
        Range r = Cell.resolveCellRange(range);
        if (r.startAddress.compareTo(r.endAddress) < 0) {
            this.startAddress = r.startAddress;
            this.endAddress = r.endAddress;
        }
        else {
            this.startAddress = r.endAddress;
            this.endAddress = r.startAddress;
        }
    }

    // ### M E T H O D S ###

    /**
     * Gets whether another range is completely enclosed by this range
     *
     * @param other Other range to check
     * @return True if the other range is completely enclosed, false if only partial overlapping or not intersecting
     */
    public boolean contains(Range other) {
        return this.startAddress.getColumn() <= other.startAddress.getColumn() &&
                this.endAddress.getColumn() >= other.endAddress.getColumn() &&
                this.startAddress.getRow() <= other.startAddress.getRow() &&
                this.endAddress.getRow() >= other.endAddress.getRow();
    }

    /**
     * Determines whether an address is within this range
     *
     * @param address Address to check
     * @return True if the address is part of this range, otherwise false
     */
    public boolean contains(Address address) {
        return address.getColumn() >= this.startAddress.getColumn() &&
                address.getColumn() <= this.endAddress.getColumn() &&
                address.getRow() >= this.startAddress.getRow() &&
                address.getRow() <= this.endAddress.getRow();
    }

    /**
     * Determines whether the passed range overlaps with this range
     *
     * @param other Range to check for overlapping
     * @return True if overlapping, otherwise false
     */
    public boolean overlaps(Range other) {
        return !(this.endAddress.getRow() < other.startAddress.getRow() ||
                this.startAddress.getRow() > other.endAddress.getRow() ||
                this.endAddress.getColumn() < other.startAddress.getColumn() ||
                this.startAddress.getColumn() > other.endAddress.getColumn());
    }

    /**
     * Gets a list of all addresses between the start and end address
     *
     * @return List of Addresses
     */
    public List<Address> resolveEnclosedAddresses() {
        return Cell.getCellRange(this.startAddress, this.endAddress);
    }

    /**
     * Overwritten toString method
     *
     * @return Returns the range (e.g. 'A1:B12')
     */
    @Override
    public String toString() {
        return startAddress.toString() + ":" + endAddress.toString();
    }

    /**
     * Overwritten equals method
     *
     * @param o Other object to compare
     * @return True if this instance is equal to the other instance
     */
    @Override
    public boolean equals(Object o) {
        if (o == this) {
            return true;
        }
        if (!(o instanceof Range)) {
            return false;
        }
        Range range = (Range) o;
        return this.startAddress.equals(range.startAddress) && this.endAddress.equals(range.endAddress);
    }

    /**
     * Gets the hash code of the range based on its string representation. The cell types (possible $ prefix) are
     * considered
     *
     * @return Hash code of the range
     */
    @Override
    public int hashCode() {
        return this.toString().hashCode();
    }

    /**
     * Creates a (dereferenced, if applicable) deep copy of this range
     *
     * @return Copy of this range
     */
    Range copy() {
        return new Range(this.startAddress.copy(), this.endAddress.copy());
    }

}
