package ch.rabanti.nanoxlsx4j.cells;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Cell.AddressType;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class AddressTest {

    @ParameterizedTest
    @DisplayName("Constructor call with string as parameter")
    @CsvSource({
        "A1, 0, 0, DEFAULT",
        "b10, 1, 9, DEFAULT",
        "$A1, 0, 0, FIXED_COLUMN",
        "A$1048576, 0, 1048575, FIXED_ROW",
        "$xFd$1, 16383, 0, FIXED_ROW_AND_COLUMN"
    })
    public void addressConstructorTest(
            String address, int expectedColumn, int expectedRow, AddressType expectedType) {
        Address actualAddress = new Address(address);
        assertEquals(expectedRow, actualAddress.row());
        assertEquals(expectedColumn, actualAddress.column());
        assertEquals(expectedType, actualAddress.type());
    }

    @ParameterizedTest
    @DisplayName("Constructor call with row and column as parameters")
    @CsvSource({
        "0, 0, A1",
        "4, 9, E10",
        "16383, 1048575, XFD1048576",
        "2, 99, C100"
    })
    public void addressConstructorTest2(int column, int row, String expectedAddress) {
        Address actualAddress = new Address(column, row);
        assertEquals(expectedAddress, actualAddress.toString());
        assertEquals(Cell.AddressType.DEFAULT, actualAddress.type());
    }

    @ParameterizedTest
    @DisplayName("Constructor call with all parameters")
    @CsvSource({
        "0, 0, DEFAULT, A1",
        "4, 9, FIXED_COLUMN, $E10",
        "16383, 1048575, FIXED_ROW, XFD$1048576",
        "2, 99, FIXED_ROW_AND_COLUMN, $C$100"
    })
    public void addressConstructorTest3(int column, int row, AddressType type, String expectedAddress) {
        Address actualAddress = new Address(column, row, type);
        assertEquals(expectedAddress, actualAddress.toString());
    }

    @ParameterizedTest
    @DisplayName("Constructor call with string and type as parameters")
    @CsvSource({
        "A1, DEFAULT, A1",
        "A1, FIXED_COLUMN, $A1",
        "A1, FIXED_ROW, A$1",
        "A1, FIXED_ROW_AND_COLUMN, $A$1",
        "$A1, DEFAULT, A1",
        "A$1, DEFAULT, A1",
        "$A$1, DEFAULT, A1"
    })
    public void addressConstructorTest4(String address, AddressType type, String expectedAddress) {
        Address actualAddress = new Address(address, type);
        assertEquals(expectedAddress, actualAddress.toString());
    }

    @ParameterizedTest
    @DisplayName("Test of Equals() implementation")
    @CsvSource({
        "A1, A1, true",
        "A1, A2, false",
        "A1, B1, false",
        "$A1, $A1, true",
        "$A1, A1, false",
        "$A1, A$1, false",
        "$A1, $A2, false",
        "$A1, $B1, false",
        "$A$1, $A$1, true",
        "$A$1, A1, false",
        "$A$1, $A1, false",
        "$A$1, $A$2, false",
        "$A$1, $B$1, false",
        "A$1, A$1, true",
        "A$1, A1, false",
        "A$1, $A1, false",
        "A$1, $A$1, false",
        "A$1, A$2, false",
        "A$1, B$1, false"
    })
    public void addressEqualsTest(String address1, String address2, boolean expectedEquality) {
        Address currentAddress = new Address(address1);
        Address otherAddress = new Address(address2);
        boolean actualEquality = currentAddress.equals(otherAddress);
        // Java has one equals(Object) method; assigning to Object enforces that call shape.
        boolean actualEquality2 = currentAddress.equals((Object) otherAddress);
        assertEquals(expectedEquality, actualEquality);
        assertEquals(expectedEquality, actualEquality2);
        if (expectedEquality) {
            assertEquals(currentAddress, otherAddress);
            assertFalse(!currentAddress.equals(otherAddress));
        } else {
            assertNotEquals(currentAddress, otherAddress);
            assertFalse(currentAddress.equals(otherAddress));
        }
    }

    @Test
    @DisplayName("Test of Equals() implementation returning false on different types")
    public void addressEqualsTest2() {
        Address currentAddress = new Address("A1");
        String other = "test";
        assertFalse(currentAddress.equals(other));
    }

    @ParameterizedTest
    @DisplayName("Test of the GetAddress method (string output)")
    @CsvSource({
        "0, 0, DEFAULT, A1",
        "4, 9, FIXED_COLUMN, $E10",
        "16383, 1048575, FIXED_ROW, XFD$1048576",
        "2, 99, FIXED_ROW_AND_COLUMN, $C$100"
    })
    public void getAddressTest(int column, int row, AddressType type, String expectedAddress) {
        Address actualAddress = new Address(column, row, type);
        assertEquals(expectedAddress, actualAddress.getAddress());
    }

    @ParameterizedTest
    @DisplayName("Test of the GetColumn function")
    @CsvSource({
        "0, 0, DEFAULT, A",
        "5, 100, FIXED_COLUMN, F",
        "26, 100, FIXED_ROW, AA",
        "1, 5, FIXED_ROW_AND_COLUMN, B"
    })
    public void getColumnTest(int columnNumber, int rowNumber, AddressType type, String expectedColumn) {
        Address address = new Address(columnNumber, rowNumber, type);
        assertEquals(expectedColumn, address.getColumn());
    }

    @ParameterizedTest
    @DisplayName("Test of GetHashCode() implementation")
    @CsvSource({
        "A1, A1, true",
        "A1, A2, false",
        "A1, B1, false",
        "$A1, $A1, true",
        "$A1, A1, false",
        "$A1, A$1, false",
        "$A1, $A2, false",
        "$A1, $B1, false",
        "$A$1, $A$1, true",
        "$A$1, A1, false",
        "$A$1, $A1, false",
        "$A$1, $A$2, false",
        "$A$1, $B$1, false",
        "A$1, A$1, true",
        "A$1, A1, false",
        "A$1, $A1, false",
        "A$1, $A$1, false",
        "A$1, A$2, false",
        "A$1, B$1, false"
    })
    public void addressGetHashCodeTest(String address1, String address2, boolean expectedEquality) {
        Address currentAddress = new Address(address1);
        Address otherAddress = new Address(address2);
        if (expectedEquality) {
            assertEquals(currentAddress.hashCode(), otherAddress.hashCode());
        } else {
            assertNotEquals(currentAddress.hashCode(), otherAddress.hashCode());
        }
    }

    @ParameterizedTest
    @DisplayName("Fail on invalid constructor calls with an address string")
    @MethodSource("invalidAddressConstructorArguments")
    public void addressConstructorFailTest(String address, Class<? extends Exception> expectedExceptionType) {
        Exception exception = assertThrows(expectedExceptionType, () -> new Address(address));
        assertEquals(expectedExceptionType, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Fail on invalid constructor calls with column and row numbers")
    @CsvSource({
        "0, -100",
        "-100, 0",
        "-1, -1",
        "16384, 0",
        "0, 1048576"
    })
    public void addressConstructorFailTest2(int column, int row) {
        assertThrows(RangeException.class, () -> new Address(column, row, AddressType.DEFAULT));
    }

    @ParameterizedTest
    @DisplayName("Test of the CompareTo function")
    @CsvSource({
        "A1, A1, 0",
        "A10, A2, 1",
        "B2, D4, -1",
        "$X$99, X99, 0",
        "A100, A$20, 1",
        "$C$2, $D$4, -1"
    })
    public void compareToTest(String address1, String address2, int expectedResult) {
        Address address = new Address(address1);
        Address otherAddress = new Address(address2);
        int result = address.compareTo(otherAddress);
        assertEquals(expectedResult, result);
    }

    @ParameterizedTest
    @DisplayName("Test of the < address operator")
    @CsvSource({
        "A1, A1, false",
        "A10, A2, false",
        "B2, D4, true",
        "$X$99, X99, false",
        "A100, A$20, false",
        "$C$2, $D$4, true"
    })
    public void compareToTest2(String address1, String address2, boolean expectedResult) {
        Address address = new Address(address1);
        Address otherAddress = new Address(address2);
        boolean result = address.compareTo(otherAddress) < 0;
        assertEquals(expectedResult, result);
    }

    @ParameterizedTest
    @DisplayName("Test of the <= address operator")
    @CsvSource({
        "A1, A1, true",
        "A10, A2, false",
        "B2, D4, true",
        "$X$99, X99, true",
        "A100, A$20, false",
        "$C$2, $D$4, true"
    })
    public void compareToTest3(String address1, String address2, boolean expectedResult) {
        Address address = new Address(address1);
        Address otherAddress = new Address(address2);
        boolean result = address.compareTo(otherAddress) <= 0;
        assertEquals(expectedResult, result);
    }

    @ParameterizedTest
    @DisplayName("Test of the > address operator")
    @CsvSource({
        "A1, A1, false",
        "A10, A2, true",
        "B2, D4, false",
        "$X$99, X99, false",
        "A100, A$20, true",
        "$C$2, $D$4, false"
    })
    public void compareToTest4(String address1, String address2, boolean expectedResult) {
        Address address = new Address(address1);
        Address otherAddress = new Address(address2);
        boolean result = address.compareTo(otherAddress) > 0;
        assertEquals(expectedResult, result);
    }

    @ParameterizedTest
    @DisplayName("Test of the >= address operator")
    @CsvSource({
        "A1, A1, true",
        "A10, A2, true",
        "B2, D4, false",
        "$X$99, X99, true",
        "A100, A$20, true",
        "$C$2, $D$4, false"
    })
    public void compareToTest5(String address1, String address2, boolean expectedResult) {
        Address address = new Address(address1);
        Address otherAddress = new Address(address2);
        boolean result = address.compareTo(otherAddress) >= 0;
        assertEquals(expectedResult, result);
    }

    // C# ExplicitOperatorTest is omitted because Java has no user-defined conversion operators. String-construction
    // behavior is covered by addressConstructorTest and addressConstructorTest4.

    private static Stream<Arguments> invalidAddressConstructorArguments() {
        return Stream.of(
            Arguments.of(null, FormatException.class),
            Arguments.of("", FormatException.class),
            Arguments.of("$", FormatException.class),
            Arguments.of("2", FormatException.class),
            Arguments.of("$D", FormatException.class),
            Arguments.of("$2", FormatException.class),
            Arguments.of("Z", FormatException.class),
            Arguments.of("A1048577", RangeException.class),
            Arguments.of("XFE1", RangeException.class)
        );
    }
}
