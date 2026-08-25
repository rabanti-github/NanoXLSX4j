package ch.rabanti.nanoxlsx4j.cells;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Cell.AddressType;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class CellsTest {

    @ParameterizedTest
    @DisplayName("Test of the resolveCellAddress method")
    @CsvSource({
        "0, 0, DEFAULT, A1",
        "0, 0, FIXED_COLUMN, $A1",
        "0, 0, FIXED_ROW, A$1",
        "0, 0, FIXED_ROW_AND_COLUMN, $A$1",
        "5, 99, DEFAULT, F100",
        "5, 99, FIXED_COLUMN, $F100",
        "5, 99, FIXED_ROW, F$100",
        "5, 99, FIXED_ROW_AND_COLUMN, $F$100",
        "16383, 1048575, DEFAULT, XFD1048576",
        "16383, 1048575, FIXED_COLUMN, $XFD1048576",
        "16383, 1048575, FIXED_ROW, XFD$1048576",
        "16383, 1048575, FIXED_ROW_AND_COLUMN, $XFD$1048576"
    })
    public void resolveCellAddressTest(int column, int row, AddressType type, String expectedAddress) {
        String address = Cell.resolveCellAddress(column, row, type);
        assertEquals(expectedAddress, address);
    }

    @ParameterizedTest
    @DisplayName("Test of the resolveCellAddress method overload")
    @CsvSource({
        "0, 0, A1",
        "5, 99, F100",
        "16383, 1048575, XFD1048576"
    })
    public void resolveCellAddressOverloadTest(int column, int row, String expectedAddress) {
        String address = Cell.resolveCellAddress(column, row);
        assertEquals(expectedAddress, address);
    }

    @ParameterizedTest
    @DisplayName("Test of the ResolveColumnAddress method")
    @CsvSource({
        "0, A",
        "2, C",
        "16383, XFD"
    })
    public void resolveColumnAddressTest(int columnNumber, String expectedAddress) {
        String address = Cell.resolveColumnAddress(columnNumber);
        assertEquals(expectedAddress, address);
    }

    @Test()
    @DisplayName("Test of the failing ResolveColumnAddress method")
    public void resolveColumnAddressFailingTest(){
        Exception ex = assertThrows(RangeException.class, () -> Cell.resolveColumnAddress(-1));
        assertEquals(RangeException.class, ex.getClass());
        ex = assertThrows(RangeException.class, () -> Cell.resolveColumnAddress(16384));
        assertEquals(RangeException.class, ex.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the validateColumnNumber method")
    @CsvSource({
        "-1, false",
        "0, true",
        "1, true",
        "16382, true",
        "16383, true",
        "16384, false"
    })
    public void validateColumnNumberTest(int column, boolean valid) {
        if (valid) {
            assertDoesNotThrow(() -> Cell.validateColumnNumber(column));
        } else {
            assertThrows(RangeException.class, () -> Cell.validateColumnNumber(column));
        }
    }

    @ParameterizedTest
    @DisplayName("Test of the validateRowNumber method")
    @CsvSource({
        "-1, false",
        "0, true",
        "1, true",
        "1048574, true",
        "1048575, true",
        "1048576, false"
    })
    public void validateRowNumberTest(int row, boolean valid) {
        if (valid) {
            assertDoesNotThrow(() -> Cell.validateRowNumber(row));
        } else {
            assertThrows(RangeException.class, () -> Cell.validateRowNumber(row));
        }
    }
}
