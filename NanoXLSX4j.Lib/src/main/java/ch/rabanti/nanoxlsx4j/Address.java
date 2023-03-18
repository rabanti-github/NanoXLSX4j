/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import static ch.rabanti.nanoxlsx4j.Cell.resolveCellAddress;

import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

/**
 * Class representing a cell address (no getters and setters to simplify
 * handling)
 *
 * @author Raphael Stoeckli
 */
public class Address implements Comparable<Address> {

	// ### P U B L I C F I E L D S ###
	/**
	 * Column of the address (zero-based)
	 */
	public final int Column;
	/**
	 * Row of the address (zero-based)
	 */
	public final int Row;

	public final Cell.AddressType Type;

	// ### C O N S T R U C T O R S ###

	/**
	 * Constructor with with row and column as arguments. The referencing type of
	 * the address is default (e.g. 'C20')
	 *
	 * @param column
	 *            Column of the address (zero-based)
	 * @param row
	 *            Row of the address (zero-based)
	 * @throws RangeException
	 *             Thrown if the resolved address is out of range
	 */
	public Address(int column, int row) {
		this(column, row, Cell.AddressType.Default);
	}

	/**
	 * Constructor with with row, column and address type as arguments
	 *
	 * @param column
	 *            Column of the address (zero-based)
	 * @param row
	 *            Row of the address (zero-based)
	 * @param type
	 *            Referencing type of the address
	 * @throws RangeException
	 *             Thrown if the resolved address is out of range
	 */
	public Address(int column, int row, Cell.AddressType type) {
		Cell.validateRowNumber(row);
		Cell.validateColumnNumber(column);
		this.Column = column;
		this.Row = row;
		this.Type = type;
	}

	/**
	 * Constructor with address as string. If no referencing modifiers ($) are
	 * defined, the address is of referencing type default (e.g. 'C23')
	 *
	 * @param address
	 *            Address string (e.g. 'B12')
	 * @throws RangeException
	 *             Thrown if the resolved address is out of range
	 */
	public Address(String address) {
		Address a = Cell.resolveCellCoordinate(address);
		this.Column = a.Column;
		this.Row = a.Row;
		this.Type = a.Type;
	}

	/**
	 * Constructor with address as string. All referencing modifiers ($) are ignored
	 * and only the defined referencing type considered
	 *
	 * @param address
	 *            Address string (e.g. 'B12')
	 * @param type
	 *            Referencing type of the address
	 * @throws RangeException
	 *             Thrown if the resolved address is out of range
	 */
	public Address(String address, Cell.AddressType type) {
		this.Type = type;
		Address a = Cell.resolveCellCoordinate(address);
		this.Column = a.Column;
		this.Row = a.Row;
	}

	// ### M E T H O D S ###

	/**
	 * Gets the address as string
	 *
	 * @return Address as string
	 * @throws RangeException
	 *             Thrown if the column or row is out of range
	 */
	public String getAddress() {
		return resolveCellAddress(this.Column, this.Row, this.Type);
	}

	/**
	 * Gets the column address (A - XFD)
	 *
	 * @return Column address as letter(s)
	 */
	public String getColumn() {
		return Cell.resolveColumnAddress(Column);
	}

	/**
	 * Returns the address as string or "Illegal Address" in case of an exception
	 *
	 * @return Address or notification in case of an error
	 */
	@Override
	public String toString() {
		return getAddress(); // Validity already checked in method
	}

	/**
	 * Compares two addresses whether they are equal. The address type id neglected,
	 * only the row and column is considered
	 *
	 * @param o
	 *            Address to compare
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
		return this.Row == address.Row && this.Column == address.Column && this.Type == address.Type;
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
	 * Compares two addresses and returns whether the passed one is bigger or not.
	 * Bigger means: Higher column number and / or higher row number
	 *
	 * @param other
	 *            Other address to compare
	 * @return 1 If the other address is bigger, -1 if it is smaller and 0 ist it is
	 *         equal
	 */
	@Override
	public int compareTo(Address other) {
		long thisCoordinate = (long) Column * (long) Worksheet.MAX_ROW_NUMBER + Row;
		long otherCoordinate = (long) other.Column * (long) Worksheet.MAX_ROW_NUMBER + other.Row;
		return Long.compare(thisCoordinate, otherCoordinate);
	}

	/**
	 * Creates a (dereferenced, if applicable) deep copy of this address
	 *
	 * @return Copy of this range
	 */
	Address copy() {
		return new Address(this.Column, this.Row, this.Type);
	}

}