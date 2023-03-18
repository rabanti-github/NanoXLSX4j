/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.util.List;

/**
 * Class representing a cell range (no getters and setters to simplify handling)
 *
 * @author Raphael Stoeckli
 */
public class Range {
	// ### P U B L I C F I E L D S ###
	/**
	 * End address of the range
	 */
	public final Address EndAddress;
	/**
	 * Start address of the range
	 */
	public final Address StartAddress;

	// ### C O N S T R U C T O R S ###

	/**
	 * Constructor with with addresses as arguments. The addresses are automatically
	 * swapped if the start address is greater than the end address
	 *
	 * @param start
	 *            Start address of the range
	 * @param end
	 *            End address of the range
	 */
	public Range(Address start, Address end) {
		if (start.compareTo(end) < 0) {
			StartAddress = start;
			EndAddress = end;
		}
		else {
			StartAddress = end;
			EndAddress = start;
		}
	}

	/**
	 * Constructor with a range string as argument. The addresses are automatically
	 * swapped if the start address is greater than the end address
	 *
	 * @param range
	 *            Address range (e.g. 'A1:B12')
	 */
	public Range(String range) {
		Range r = Cell.resolveCellRange(range);
		if (r.StartAddress.compareTo(r.EndAddress) < 0) {
			StartAddress = r.StartAddress;
			EndAddress = r.EndAddress;
		}
		else {
			StartAddress = r.EndAddress;
			EndAddress = r.StartAddress;
		}
	}

	// ### M E T H O D S ###

	/**
	 * Gets a list of all addresses between the start and end address
	 *
	 * @return List of Addresses
	 */
	public List<Address> resolveEnclosedAddresses() {
		return Cell.getCellRange(this.StartAddress, this.EndAddress);
	}

	/**
	 * Overwritten toString method
	 *
	 * @return Returns the range (e.g. 'A1:B12')
	 */
	@Override
	public String toString() {
		return StartAddress.toString() + ":" + EndAddress.toString();
	}

	/**
	 * Overwritten equals method
	 *
	 * @param o
	 *            Other object to compare
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
		return this.StartAddress.equals(range.StartAddress) && this.EndAddress.equals(range.EndAddress);
	}

	/**
	 * Gets the hash code of the range based on its string representation. The cell
	 * types (possible $ prefix) are considered
	 * 
	 * @return Hash code of the address
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
		return new Range(this.StartAddress.copy(), this.EndAddress.copy());
	}
}