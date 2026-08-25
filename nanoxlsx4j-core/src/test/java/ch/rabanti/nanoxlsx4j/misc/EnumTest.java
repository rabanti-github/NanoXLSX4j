/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.enums.Errors;
import ch.rabanti.nanoxlsx4j.enums.FormulaError;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class EnumTest {

    @DisplayName("Test parsing formula errors")
    @ParameterizedTest
    @CsvSource({
        "'#NULL!', NULL",
        "'#DIV/0!', DIVISION_BY_ZERO",
        "'#VALUE!', VALUE",
        "'#REF!', REFERENCE",
        "'#NAME?', NAME",
        "'#NUM!', NUMBER",
        "'#N/A', NOT_AVAILABLE",
        "'#GETTING_DATA', GETTING_DATA"
    })
    public void tryParseFormulaErrorTest(String value, FormulaError expected) {
        Optional<FormulaError> result = Errors.tryParseFormulaError(value);
        boolean success = result.isPresent();
        FormulaError error = result.orElse(FormulaError.UNKNOWN_ERROR);

        assertTrue(success);
        assertEquals(expected, error);
    }

    @DisplayName("Test parsing invalid formula errors")
    @ParameterizedTest
    @NullSource
    @ValueSource(strings = {"", " ", "#NAME? ", "#name?", "#NAME", "NAME?", "#UNKNOWN!"})
    public void tryParseInvalidFormulaErrorTest(String value) {
        Optional<FormulaError> result = Errors.tryParseFormulaError(value);
        boolean success = result.isPresent();
        FormulaError error = result.orElse(FormulaError.UNKNOWN_ERROR);

        assertFalse(success);
        assertEquals(FormulaError.UNKNOWN_ERROR, error);
    }

    @DisplayName("Test conversion of formula errors to strings")
    @ParameterizedTest
    @CsvSource({
        "NULL, '#NULL!'",
        "DIVISION_BY_ZERO, '#DIV/0!'",
        "VALUE, '#VALUE!'",
        "REFERENCE, '#REF!'",
        "NAME, '#NAME?'",
        "NUMBER, '#NUM!'",
        "NOT_AVAILABLE, '#N/A'",
        "GETTING_DATA, '#GETTING_DATA'"
    })
    public void formulaErrorToStringTest(FormulaError error, String expected) {
        String result = Errors.formulaErrorToString(error);

        assertEquals(expected, result);
    }

    @DisplayName("Test conversion of invalid formula errors to strings")
    @ParameterizedTest
    // The C# value (FormulaError)999 is omitted because Java enums cannot represent undefined numeric values.
    @EnumSource(value = FormulaError.class, names = {"NO_ERROR", "UNKNOWN_ERROR"})
    public void formulaErrorToStringFailTest(FormulaError error) {
        assertThrows(FormatException.class, () -> Errors.formulaErrorToString(error));
    }
}
