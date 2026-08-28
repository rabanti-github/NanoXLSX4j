/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.enums;

import java.util.Optional;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

/** Provides conversions for formula errors. */
public final class Errors {
    private Errors() {
        // Do not instantiate
    }

    /**
     * Tries to parse a formula error from its OOXML representation.
     *
     * @param value string to parse
     * @return the parsed error, or an empty value if the string is not an OOXML formula error
     */
    public static Optional<FormulaError> tryParseFormulaError(String value) {
        if (value == null) {
            return Optional.empty();
        }

        return switch (value) {
            case "#NULL!" -> Optional.of(FormulaError.NULL);
            case "#DIV/0!" -> Optional.of(FormulaError.DIVISION_BY_ZERO);
            case "#VALUE!" -> Optional.of(FormulaError.VALUE);
            case "#REF!" -> Optional.of(FormulaError.REFERENCE);
            case "#NAME?" -> Optional.of(FormulaError.NAME);
            case "#NUM!" -> Optional.of(FormulaError.NUMBER);
            case "#N/A" -> Optional.of(FormulaError.NOT_AVAILABLE);
            case "#GETTING_DATA" -> Optional.of(FormulaError.GETTING_DATA);
            default -> Optional.empty();
        };
    }

    /**
     * Returns the OOXML-conformant error expression as a string.
     *
     * @param error enum value
     * @return OOXML internal error expression as a string
     * @throws FormatException if an invalid value such as {@link FormulaError#NO_ERROR} is passed
     */
    public static String formulaErrorToString(FormulaError error) throws FormatException {
        if (error == null) {
            throw new FormatException("An invalid error type 'null' was specified");
        }

        return switch (error) {
            case NULL -> "#NULL!";
            case DIVISION_BY_ZERO -> "#DIV/0!";
            case VALUE -> "#VALUE!";
            case REFERENCE -> "#REF!";
            case NAME -> "#NAME?";
            case NUMBER -> "#NUM!";
            case NOT_AVAILABLE -> "#N/A";
            case GETTING_DATA -> "#GETTING_DATA";
            default -> throw new FormatException("An invalid error type '" + error + "' was specified");
        };
    }
}
