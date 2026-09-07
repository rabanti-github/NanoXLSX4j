/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ComparatorsTest {

    @ParameterizedTest
    @DisplayName("Test of the comparator function compareSecureStrings")
    @CsvSource(value = {
            "''|''|true",
            "' '|' '|true",
            "a|a|true",
            "12345678|12345678|true",
            "@à#|@à#|true",
            "a|A|false",
            "''|' '|false",
            "123|1234|false",
            "...|.,.|false",
            "NULL|NULL|true",
            "NULL|''|true",
            "NULL|' '|false",
            "NULL|ABC|false"
    }, delimiter = '|', nullValues = "NULL")
    public void compareSecureStringsTest(String plainText1, String plainText2, boolean expectedEqual) {
        char[] secureString1 = getSecureString(plainText1);
        char[] secureString2 = getSecureString(plainText2);
        boolean isEqual = Comparators.compareSecureStrings(secureString1, secureString2);
        assertEquals(expectedEqual, isEqual);
        isEqual = Comparators.compareSecureStrings(secureString2, secureString1); // reverse
        assertEquals(expectedEqual, isEqual);
    }

    @ParameterizedTest
    @DisplayName("Test of the comparator function compareDimensions")
    @CsvSource(value = {
            "15.2|15.3|-1",
            "15.3|15.2|1",
            "0.0002|0.0003|-1",
            "0.0003|0.0002|1",
            "0.0002|0.0002|0",
            "1|2|-1",
            "2|1|1",
            "-1|2|-1",
            "-2|-1|-1",
            "-1|-2|1",
            "0|0|0",
            "NULL|15.3|-1",
            "15.3|NULL|1",
            "NULL|NULL|0"
    }, delimiter = '|', nullValues = "NULL")
    public void compareDimensionsTest(Float dimension1, Float dimension2, int expectedResult) {
        assertEquals(expectedResult, Comparators.compareDimensions(dimension1, dimension2));
    }

    @ParameterizedTest
    @DisplayName("Test of isZero comparator (double)")
    @CsvSource(value = {
            "0.0|true",
            "-0.0|true",
            "1e-15|true",
            "-1e-15|true",
            "1e-13|true",
            "-1e-13|true",
            "1e-11|false",
            "-1e-11|false",
            "1.0|false",
            "-1.0|false",
            "4.9e-324|true",
            "NaN|false",
            "Infinity|false",
            "-Infinity|false"
    }, delimiter = '|')
    public void isZeroDoubleTest(double value, boolean expected) {
        assertEquals(expected, Comparators.isZero(value));
    }

    @ParameterizedTest
    @DisplayName("Test of isZero comparator (float)")
    @CsvSource(value = {
            "0.0|true",
            "-0.0|true",
            "1e-8|true",
            "-1e-8|true",
            "1e-6|true",
            "-1e-6|true",
            "1e-5|false",
            "-1e-5|false",
            "1.0|false",
            "-1.0|false",
            "1.4e-45|true",
            "NaN|false",
            "Infinity|false",
            "-Infinity|false"
    }, delimiter = '|')
    public void isZeroFloatTest(float value, boolean expected) {
        assertEquals(expected, Comparators.isZero(value));
    }

    private static char[] getSecureString(String plainText) {
        if (plainText == null) {
            return null;
        }
        return plainText.toCharArray();
    }
}
