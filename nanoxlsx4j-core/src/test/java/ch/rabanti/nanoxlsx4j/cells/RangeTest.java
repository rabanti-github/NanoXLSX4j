/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.cells;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Range;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.List;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

public class RangeTest {

    @ParameterizedTest
    @DisplayName("Test of the Range constructor with start and end address")
    @CsvSource({
        "A1, A1, A1:A1",
        "A1, C4, A1:C4",
        "C3, A1, A1:C3",
        "$A1, $A$2, $A1:$A$2",
        "A$1, C$4, A$1:C$4",
        "$C$3, $A1, $A1:$C$3"
    })
    public void constructorTest(String startAddress, String endAddress, String expectedRange) {
        Address start = new Address(startAddress);
        Address end = new Address(endAddress);
        Range range = new Range(start, end);
        assertEquals(expectedRange, range.toString());
    }

    @ParameterizedTest
    @DisplayName("Test of the Range constructor with range expression string")
    @CsvSource({
        "A1:A1, A1:A1",
        "c2:C3, C2:C3",
        "$A1:$F10, $A1:$F10",
        "$r$1:$b$2, $B$2:$R$1"
    })
    public void constructorTest2(String rangeExpression, String expectedRange) {
        Range range = new Range(rangeExpression);
        assertEquals(expectedRange, range.toString());
    }

    @ParameterizedTest
    @DisplayName("Test of the Range constructor with column and row numbers")
    @CsvSource({
        "0, 0, 0, 0, A1:A1",
        "0, 0, 1, 1, A1:B2",
        "1, 1, 0, 0, A1:B2"
    })
    public void constructorTest3(
            int startColumn, int startRow, int endColumn, int endRow, String expectedRange) {
        Range range = new Range(startColumn, startRow, endColumn, endRow);
        assertEquals(expectedRange, range.toString());
    }

    @ParameterizedTest
    @DisplayName("Test of the ResolveEnclosedAddressesTest method")
    @CsvSource(delimiter = '|', value = {
        "A1:A1 | A1",
        "A1:A4 | A1,A2,A3,A4",
        "A1:B3 | A1,A2,A3,B1,B2,B3",
        "B3:A2 | A2,A3,B2,B3"
    })
    public void resolveEnclosedAddressesTest(String rangeExpression, String expectedAddresses) {
        Range range = new Range(rangeExpression);
        List<Address> addresses = range.resolveEnclosedAddresses();
        assertEquals(expectedAddresses,
            addresses.stream().map(Address::toString).collect(Collectors.joining(",")));
    }

    @ParameterizedTest
    @DisplayName("Test of the Contains method on addresses")
    @CsvSource({
        "A1:A1, A1, true",
        "B2:F5, C3, true",
        "B2:F5, F5, true",
        "B2:F5, B2, true",
        "B2:F5, B5, true",
        "B2:F5, F2, true",
        "B2:B2, B1, false",
        "B2:F5, F6, false"
    })
    public void containsTest(String rangeExpression, String givenAddress, boolean expectedResult) {
        Range range = new Range(rangeExpression);
        Address address = new Address(givenAddress);
        boolean contains = range.contains(address);
        assertEquals(contains, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the Contains method on ranges")
    @CsvSource({
        "A1:A1, A1:A1, true",
        "B2:F5, C3:C3, true",
        "B2:F5, B2:F5, true",
        "B2:F5, B2:C3, true",
        "B2:F5, E4:F5, true",
        "B2:F5, E2:F3, true",
        "B2:F5, B4:C5, true",
        "B2:F5, B1:C3, false",
        "B2:F5, E2:G3, false",
        "B2:F5, B5:B6, false",
        "B2:F5, E4:G6, false",
        "B2:F5, A1:A2, false",
        "B2:F5, G1:H2, false",
        "B2:F5, A6:B8, false",
        "B2:F5, E6:G7, false",
        "B2:B2, B1:B1, false",
        "B2:F5, H3:F6, false",
        "B2:F5, A1:G8, false"
    })
    public void containsTest2(String rangeExpression, String givenRange, boolean expectedResult) {
        Range range = new Range(rangeExpression);
        Range range2 = new Range(givenRange);
        boolean contains = range.contains(range2);
        assertEquals(contains, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the Overlaps method")
    @CsvSource({
        "A1:A1, A1:A1, true",
        "B2:F5, C3:C3, true",
        "B2:F5, B2:F5, true",
        "B2:F5, B2:C3, true",
        "B2:F5, E4:F5, true",
        "B2:F5, E2:F3, true",
        "B2:F5, B4:C5, true",
        "B2:F5, A1:G8, true",
        "B2:F5, B1:C3, true",
        "B2:F5, E2:G3, true",
        "B2:F5, B5:B6, true",
        "B2:F5, E4:G6, true",
        "B2:F5, A1:A2, false",
        "B2:F5, G1:H2, false",
        "B2:F5, A6:B8, false",
        "B2:F5, E6:G7, false",
        "B2:B2, B1:B1, false",
        "B2:F5, H3:F6, false"
    })
    public void overlapsTest(String rangeExpression, String givenRange, boolean expectedResult) {
        Range range = new Range(rangeExpression);
        Range range2 = new Range(givenRange);
        boolean contains = range.Overlaps(range2);
        assertEquals(contains, expectedResult);
    }

    // C# ImplicitOperatorTest is omitted because Java has no user-defined conversion operators. Range string parsing
    // is covered by constructorTest2 through the corresponding String constructor.

    @ParameterizedTest
    @DisplayName("Test of the Equals method")
    @CsvSource({
        "A1:A1, A1:A1, true",
        "A1:A4, A$1:A$4, false",
        "A1:B3, A1:B4, false",
        "B3:A2, A2:B3, true",
        "B$3:A2, A2:B$3, true"
    })
    public void equalsTest(String rangeExpression1, String rangeExpression2, boolean expectedEquality) {
        Range range1 = new Range(rangeExpression1);
        Range range2 = new Range(rangeExpression2);
        boolean result = range1.equals(range2);
        assertEquals(expectedEquality, result);
        if (expectedEquality) {
            assertEquals(range1, range2);
        } else {
            assertNotEquals(range1, range2);
        }
    }

    @Test
    @DisplayName("Test of the Equals method returning false on invalid values")
    public void equalsTest2() {
        Range range1 = new Range("A1:A7");
        boolean result = range1.equals(null);
        assertFalse(result);
        result = range1.equals("Wrong type");
        assertFalse(result);
    }

    @ParameterizedTest
    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    @CsvSource({
        "A1:A1, A1:A1, true",
        "A1:A4, A$1:A$4, false",
        "A1:B3, A1:B4, false",
        "B3:A2, A2:B3, true",
        "B$3:A2, A2:B$3, true"
    })
    public void getHashCodeTest(String rangeExpression1, String rangeExpression2, boolean expectedEquality) {
        Range range1 = new Range(rangeExpression1);
        Range range2 = new Range(rangeExpression2);
        if (expectedEquality) {
            assertEquals(range1.hashCode(), range2.hashCode());
        } else {
            assertNotEquals(range1.hashCode(), range2.hashCode());
        }
    }
}
