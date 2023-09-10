package ch.rabanti.nanoxlsx4j.cells;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class AddressTest {

    @DisplayName("Constructor call with string as parameter")
    @ParameterizedTest(name = "Given address string:{0} should lead to the column: {1}, the row: {2} and the type: {3}")
    @CsvSource(
            {
                    "A1, 0, 0, Default",
                    "b10, 1, 9, Default",
                    "$A1, 0, 0, FixedColumn",
                    "A$1048576, 0, 1048575, FixedRow",
                    "$xFd$1, 16383, 0, FixedRowAndColumn",}
    )
    void addressConstructorTest(String address, int expectedColumn, int expectedRow, Cell.AddressType expectedType) {
        Address actualAddress = new Address(address);
        assertEquals(expectedRow, actualAddress.Row);
        assertEquals(expectedColumn, actualAddress.Column);
        assertEquals(expectedType, actualAddress.Type);
    }

    @DisplayName("Constructor call with row and column as parameters")
    @ParameterizedTest(name = "Given column:{0} and row: {1} should lead to the address string: {2}")
    @CsvSource(
            {
                    "0, 0, A1",
                    "4, 9, E10",
                    "16383, 1048575, XFD1048576",
                    "2, 99, C100",}
    )
    void addressConstructorTest2(int column, int row, String expectedAddress) {
        Address actualAddress = new Address(column, row);
        assertEquals(expectedAddress, actualAddress.toString());
        assertEquals(Cell.AddressType.Default, actualAddress.Type);
    }

    @DisplayName("Constructor call with all parameters")
    @ParameterizedTest(name = "Given column:{0}, row: {1} and type: {2} should lead to the address string: {3}")
    @CsvSource(
            {
                    "0, 0, Default, A1",
                    "4, 9, FixedColumn, $E10",
                    "16383, 1048575, FixedRow, XFD$1048576",
                    "2, 99, FixedRowAndColumn, $C$100",}
    )
    void addressConstructorTest3(int column, int row, Cell.AddressType type, String expectedAddress) {
        Address actualAddress = new Address(column, row, type);
        assertEquals(expectedAddress, actualAddress.toString());
    }

    @DisplayName("Constructor call with string and type as parameters")
    @ParameterizedTest(name = "Given address :{0} and type: {2} should lead to the address string: {3}")
    @CsvSource(
            {
                    "A1, Default, A1",
                    "A1, FixedColumn, $A1",
                    "A1, FixedRow, A$1",
                    "A1, FixedRowAndColumn, $A$1",
                    "$A1, Default, A1",
                    "A$1, Default, A1",
                    "$A$1, Default, A1",}
    )
    void addressConstructorTest4(String address, Cell.AddressType type, String expectedAddress) {
        Address actualAddress = new Address(address, type);
        assertEquals(expectedAddress, actualAddress.toString());
    }

    @DisplayName("Test of equals() implementation")
    @ParameterizedTest(name = "Giver address {0} and address {1} should lead to {2} ig an equals check is performed")
    @CsvSource(
            {
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
                    "A$1, B$1, false",}
    )
    void addressEqualsTest(String address1, String address2, boolean expectedEquality) {
        Address currentAddress = new Address(address1);
        Address otherAddress = new Address(address2);
        boolean actualEquality = currentAddress.equals(otherAddress);
        assertEquals(expectedEquality, actualEquality);
        assertEquals(currentAddress, currentAddress); // Self-test
    }

    @DisplayName("Test of equals() if the other object is invalid")
    @Test
    void addressEqualsTest2() {
        Address address = new Address("A1");
        assertNotEquals(null, address);
        assertNotEquals("Incompatible type", address);
    }

    @DisplayName("Test of hashCode() implementation")
    @ParameterizedTest(name = "Giver address {0} and address {1} should lead to {2} ig an equals check is performed")
    @CsvSource(
            {
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
                    "A$1, B$1, false",}
    )
    void addressHashCodeTest(String address1, String address2, boolean expectedEquality) {
        Address currentAddress = new Address(address1);
        Address otherAddress = new Address(address2);
        if (expectedEquality) {
            assertEquals(currentAddress.hashCode(), otherAddress.hashCode());
        }
        else {
            assertNotEquals(currentAddress.hashCode(), otherAddress.hashCode());
        }
    }

    @DisplayName("Test of the GetAddress method (String output)")
    @ParameterizedTest(name = "Given column: {0}, row: {1} and type: {2} should lead to the address string {3}")
    @CsvSource(
            {
                    "0, 0, Default, A1",
                    "4, 9, FixedColumn, $E10",
                    "16383, 1048575, FixedRow, XFD$1048576",
                    "2, 99, FixedRowAndColumn, $C$100",}
    )
    void getAddressTest(int column, int row, Cell.AddressType type, String expectedAddress) {
        Address actualAddress = new Address(column, row, type);
        assertEquals(expectedAddress, actualAddress.getAddress());
    }

    @DisplayName("Test of the GetColumn function")
    @ParameterizedTest(name = "Given column: {0}, row: {1} and type {2} should lead to string: {3}")
    @CsvSource(
            {
                    "0,0, Default, A",
                    "5, 100, FixedColumn, F",
                    "26, 100, FixedRow, AA",
                    "1, 5, FixedRowAndColumn, B",}
    )
    void getColumnTest(int columnNumber, int rowNumber, Cell.AddressType type, String expectedColumn) {
        Address address = new Address(columnNumber, rowNumber, type);
        assertEquals(expectedColumn, address.getColumn());
    }

    // FAILING TESTS
    // Tests which expects an exception

    @DisplayName("Fail on invalid constructor calls with an address string")
    @ParameterizedTest(name = "Given address string {0} should lead to an exception of type {1}")
    @CsvSource(
            {
                    ", FormatException",
                    "'', FormatException",
                    "$, FormatException",
                    "2, FormatException",
                    "$D, FormatException",
                    "$2, FormatException",
                    "Z, FormatException",
                    "A1048577, RangeException",
                    "XFE1, RangeException",}
    )
    void addressConstructorFailTest(String address, String expectedExceptionType) {
        Exception ex;
        if (expectedExceptionType.equals(FormatException.class.getSimpleName())) {
            // Common malformed addresses
            assertThrows(FormatException.class, () -> {
                new Address(address);
            });
        }
        else {
            // Out of range addresses
            assertThrows(RangeException.class, () -> {
                new Address(address);
            });
        }
    }

    @DisplayName("Fail on invalid constructor calls with column and row numbers")
    @ParameterizedTest(name = "Given column: {0} and row: {1} should lead to an exception")
    @CsvSource(
            {
                    "0, -100",
                    "-100, 0",
                    "-1, -1",
                    "16384, 0",
                    "0, 1048576",}
    )
    void addressConstructorFailTest2(int column, int row) {
        assertThrows(RangeException.class, () -> {
            new Address(column, row, Cell.AddressType.Default);
        });
    }

    @DisplayName("Test of the compareTo function")
    @ParameterizedTest(name = "Given address {0} compares to {1} should lead to {2}")
    @CsvSource(
            {
                    "A1, A1, 0",
                    "A10, A2, 1",
                    "B2, D4, -1",
                    "$X$99, X99, 0",
                    "A100, A$20, 1",
                    "$C$2, $D$4, -1",}
    )
    void compareToTest(String address1, String address2, int expectedResult) {
        Address address = new Address(address1);
        Address otherAddress = new Address(address2);
        int result = address.compareTo(otherAddress);
        assertEquals(expectedResult, result);
    }

}
