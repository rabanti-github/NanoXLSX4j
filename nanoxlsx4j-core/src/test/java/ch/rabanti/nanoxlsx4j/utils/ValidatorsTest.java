/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ValidatorsTest {

    @ParameterizedTest
    @DisplayName("Test of the successful Validator function validateColor")
    @CsvSource(value = {
            "000000, false, true",
            "000000, false, false",
            "00AACC, false, true",
            "00AACC, false, false",
            "FFFFFF, false, true",
            "FFFFFF, false, false",
            "00000000, true, true",
            "00000000, true, false",
            "FF000000, true, true",
            "FF000000, true, false",
            "00AACC00, true, true",
            "00AACC00, true, false",
            "'', true, true",
            "'', false, true",
            "NULL, true, true",
            "NULL, false, true"
    }, nullValues = "NULL")
    public void validateColorTest(String givenHexCode, boolean givenUseAlpha, boolean givenAllowEmpty) {
        Validators.validateColor(givenHexCode, givenUseAlpha, givenAllowEmpty);
        assertTrue(true);
    }

    @ParameterizedTest
    @DisplayName("Test of the failing Validator function validateColor")
    @CsvSource(value = {
            "000000, true, true",
            "FFFFFF, true, true",
            "0ACFD, true, true",
            "00000000, false, true",
            "00FFFFFF, false, true",
            "000ACFD, false, true",
            "FF000000, false, true",
            "FFFFFFFF, false, true",
            "FF0ACFD, false, true",
            "AA, false, true",
            "CCC, false, true",
            "DDDD, false, true",
            "001122, true, true",
            "X, false, true",
            "AAX022, false, true",
            "' ', false, true",
            "'0 0000', false, true",
            "'', false, false",
            "NULL, false, false"
    }, nullValues = "NULL")
    public void validateColorFailTest(String givenHexCode, boolean givenUseAlpha, boolean givenAllowEmpty) {
        assertThrows(StyleException.class,
                () -> Validators.validateColor(givenHexCode, givenUseAlpha, givenAllowEmpty));
    }

    @ParameterizedTest
    @DisplayName("Test of the successful Validator function validateGenericColor")
    @CsvSource(value = {
            "000000, true",
            "000000, false",
            "00AACC, true",
            "00AACC, false",
            "FFFFFF, true",
            "FFFFFF, false",
            "00000000, true",
            "00000000, false",
            "FF000000, true",
            "FF000000, false",
            "00AACC00, true",
            "00AACC00, false",
            "'', true",
            "NULL, true"
    }, nullValues = "NULL")
    public void validateGenericColorTest(String givenHexCode, boolean givenAllowEmpty) {
        Validators.validateGenericColor(givenHexCode, givenAllowEmpty);
        assertTrue(true);
    }

    @ParameterizedTest
    @DisplayName("Test of the failing Validator function validateGenericColor")
    @CsvSource(value = {
            "0ACFD, true",
            "000ACFD, true",
            "FF0ACFD, true",
            "AA, true",
            "CCC, true",
            "DDDD, true",
            "X, true",
            "AAX022, true",
            "' ', true",
            "'0 0000', true",
            "'', false",
            "NULL, false"
    }, nullValues = "NULL")
    public void validateGenericColorFailTest(String givenHexCode, boolean givenAllowEmpty) {
        assertThrows(StyleException.class, () -> Validators.validateGenericColor(givenHexCode, givenAllowEmpty));
    }

    @ParameterizedTest
    @DisplayName("Test of the successful Validator function validateCellAddressExpression")
    @CsvSource(value = {
            "A1, ANY",
            "$XFD$1048576, ANY",
            "A1:B2, ANY",
            "A1, SINGLE_ADDRESS",
            "$A$1, SINGLE_ADDRESS",
            "A1:B2, RANGE",
            "A1:A1, RANGE",
            "NULL, INVALID",
            "'', INVALID",
            "' ', INVALID",
            "A1:B2:C3, INVALID",
            "$A$$1, INVALID",
            "世界1, INVALID",
            "A0, INVALID"
    }, nullValues = "NULL")
    public void validateCellAddressExpressionTest(String givenExpression, Cell.AddressScope givenScope) {
        Validators.validateCellAddressExpression(givenExpression, givenScope);
        assertTrue(true);
    }

    @ParameterizedTest
    @DisplayName("Test of the Validator function validateCellAddressExpression overload")
    @CsvSource(value = {
            "A1, true",
            "$XFD$1048576, true",
            "A1:B2, true",
            "NULL, false",
            "'', false",
            "A1:B2:C3, false"
    }, nullValues = "NULL")
    public void validateCellAddressExpressionOverloadTest(String givenExpression, boolean expectedValid) {
        if (expectedValid) {
            Validators.validateCellAddressExpression(givenExpression);
            assertTrue(true);
        } else {
            assertThrows(FormatException.class, () -> Validators.validateCellAddressExpression(givenExpression));
        }
    }

    @ParameterizedTest
    @DisplayName("Test of the failing Validator function validateCellAddressExpression")
    @CsvSource(value = {
            "NULL, ANY",
            "'', ANY",
            "' ', ANY",
            "世界1, ANY",
            "A0, ANY",
            "XFE1, ANY",
            "A1048577, ANY",
            "A1:B2:C3, ANY",
            "NULL, SINGLE_ADDRESS",
            "A1:B2, SINGLE_ADDRESS",
            "A1:A1, SINGLE_ADDRESS",
            "NULL, RANGE",
            "A1, RANGE",
            "A1:B2:C3, RANGE",
            "$A$$1, SINGLE_ADDRESS"
    }, nullValues = "NULL")
    public void validateCellAddressExpressionFailTest(String givenExpression, Cell.AddressScope givenScope) {
        FormatException exception = assertThrows(FormatException.class,
                () -> Validators.validateCellAddressExpression(givenExpression, givenScope));
        assertNotNull(exception.getCause());
    }

    @ParameterizedTest
    @DisplayName("Test of the inverted Validator function ValidateCellAddressExpression")
    @ValueSource(strings = {"A1", "$XFD$1048576", "A1:B2"})
    public void validateCellAddressExpressionInvalidScopeFailTest(String givenExpression) {
        FormatException exception = assertThrows(FormatException.class,
                () -> Validators.validateCellAddressExpression(givenExpression, Cell.AddressScope.INVALID));
        assertNull(exception.getCause());
        // TODO fix this term if the naming changes in ValidateCellAddressExpression
        assertEquals("The passed expression is valid cell address or range, but the validation was explicitly inverted",
                exception.getMessage());
    }

    @ParameterizedTest
    @DisplayName("Test of the ValidateWorksheetName function")
    @CsvSource(value = {
            "1, true",
            "test, true",
            "test-test, true",
            "$$$, true",
            "'a b', true",
            "'a\tb', true",
            "-------------------------------, true",
            "'', false",
            "NULL, false",
            "'a[b', false",
            "'a]b', false",
            "'a*b', false",
            "'a?b', false",
            "'a/b', false",
            "'a\\b', false",
            "--------------------------------, false"
    }, nullValues = "NULL")
    public void setSheetNameTest(String name, boolean expectedValid) {
        Worksheet worksheet = new Worksheet(null, 0, null);
        assertNull(worksheet.getSheetName());
        if (expectedValid) {
            Validators.validateWorksheetName(name);
            assertTrue(true);
        } else {
            assertThrows(FormatException.class, () -> Validators.validateWorksheetName(name));
        }
    }
}
