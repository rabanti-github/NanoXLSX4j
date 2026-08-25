/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.enums;

/** Errors that can occur in formulas and functions. */
public enum FormulaError {
    /** Default value if no error has occurred (not an actual error type). */
    NO_ERROR,
    /** Value if a not yet defined value has occurred (not an actual error type). */
    UNKNOWN_ERROR,
    /** Indicates that two areas are required to intersect, but do not. */
    NULL,
    /** Indicates that any number (including zero) or any error code is divided by zero. */
    DIVISION_BY_ZERO,
    /**
     * Indicates that an incompatible type argument is passed to a function, or an incompatible type operand is used
     * with an operator.
     */
    VALUE,
    /** Indicates that a cell reference cannot be evaluated. */
    REFERENCE,
    /** Indicates that what looks like a name is used, but no such name has been defined. */
    NAME,
    /**
     * Indicates that an argument to a function has a compatible type, but has a value that is outside the domain over
     * which that function is defined.
     * <p>This is known as a domain error. In contrast to {@link #VALUE}, a formula or function may look valid, but
     * tries to handle an invalid value of the valid type. Example: {@code ATANH(50)}</p>
     */
    NUMBER,
    /**
     * Indicates that a designated value is not available.
     * <p>This can happen if a formula requires two arrays of the same length, but the provided ones have different
     * lengths. Example: {@code SUMX2MY2(A1:A3;B1:B4)}</p>
     */
    NOT_AVAILABLE,
    /**
     * Indicates that a cell reference cannot be evaluated because its value has not been retrieved or calculated.
     * <p>In contrast to {@link #NOT_AVAILABLE}, the value will eventually be available, for example from an external
     * source.</p>
     */
    GETTING_DATA
}
