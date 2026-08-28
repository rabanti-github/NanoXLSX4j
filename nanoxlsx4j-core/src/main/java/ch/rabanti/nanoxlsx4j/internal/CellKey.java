/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal;

import ch.rabanti.nanoxlsx4j.Cell;

/**
 * Compact, hash-efficient key for the internal cell dictionary. Uses integer (col, row) coordinates instead of rendered
 * address strings, eliminating string allocation on every cell insert/lookup.
 * <p>This class is for internal use only.</p>
 */
public final class CellKey {

    /** Column number of the key. */
    public final int column;

    /** Row number of the key. */
    public final int row;

    /**
     * Constructor with all parameters
     *
     * @param column Column number
     * @param row    Row number
     */
    public CellKey(int column, int row) {
        this.column = column;
        this.row = row;
    }

    /**
     * Returns whether two instances of CellKey are the same
     *
     * @param other The reference CellKey with which to compare
     * @return True if this instance and the other are the same
     */
    public final boolean equals(CellKey other) {
        return column == other.column && row == other.row;
    }

    /**
     * Returns whether two instances are the same
     *
     * @param o The reference object with which to compare
     * @return True if this instance and the other are the same
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof CellKey cellKey)) {
            return false;
        }

        return column == cellKey.column && row == cellKey.row;
    }

    // Excel max: 16 384 columns (14 bits) × 1 048 576 rows (20 bits) — fits cleanly in 34 bits,
    // so a simple multiply+XOR gives a collision-free hash across the valid address space.

    /**
     * Gets the hash code of the cell key
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return (row * 16384) ^ column;
    }

    /**
     * Returns the String representation of the cell key
     *
     * @return Cell address
     */
    @Override
    public String toString() {
        return Cell.resolveCellAddress(column, row);
    }
}
