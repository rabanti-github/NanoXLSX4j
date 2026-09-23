package ch.rabanti.nanoxlsx4j.cells;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

class ColumnTest {

    @ParameterizedTest
    @DisplayName("Test of the getColumnAddress and setColumnAddress of Column, as well as proper modification")
    @CsvSource(
            {
                    "A, A, B, B",
                    "a, A, b, B",
                    "AAB, AAB, A, A",
                    "a, A, XFD, XFD"
            }
    )
    void columnAddressTest(
            String initialValue, String expectedValue, String changedValue, String expectedChangedValue) {
        Column column = new Column(initialValue);
        assertEquals(expectedValue, column.getColumnAddress());
        column.setColumnAddress(changedValue);
        assertEquals(expectedChangedValue, column.getColumnAddress());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing setColumnAddress of Column")
    @NullSource // Provides null as value
    @ValueSource(
            strings = {
                    "",
                    "4",
                    "-",
                    ".",
                    "$",
                    "XFE"}
    )
    void columnAddressTest2(String value) {
        Column column = new Column("A");
        assertThrows(RangeException.class, () -> column.setColumnAddress(value));
    }

    @Test
    @DisplayName("Test of the hasAutoFilter and setAutoFilter of Column, as well as the constructor and proper " +
            "modification")
    void hasAutoFilterTest() {
        Column column = new Column("A");
        assertFalse(column.hasAutoFilter());
        column.setAutoFilter(true);
        assertTrue(column.hasAutoFilter());
    }

    @Test
    @DisplayName("Test of the isHidden and setHidden of Column, as well as proper modification")
    void isHiddenTest() {
        Column column = new Column("A");
        assertFalse(column.isHidden());
        column.setHidden(true);
        assertTrue(column.isHidden());
    }

    @ParameterizedTest
    @DisplayName("Test of the getNumber and setNumber of Column, as well as the constructor and proper modification")
    @CsvSource(
            {
                    "0, 0, 1, 1",
                    "999, 999, 5, 5",
                    "0, 0, 16383, 16383"
            }
    )
    void numberTest(int initialValue, int expectedValue, int changedValue, int expectedChangedValue) {
        Column column = new Column(initialValue);
        assertEquals(expectedValue, column.getNumber());
        column.setNumber(changedValue);
        assertEquals(expectedChangedValue, column.getNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing setNumber of Column")
    @ValueSource(
            ints = {
                    -1,
                    16384}
    )
    void numberTest2(int value) {
        Column column = new Column(2);
        assertThrows(RangeException.class, () -> column.setNumber(value));
    }

    @ParameterizedTest
    @DisplayName("Test of the getWidth and setWidth of Column, as well as proper modification")
    @CsvSource(
            {
                    "15, 15",
                    "11.1, 11.1",
                    "0, 0",
                    "255, 255"}
    )
    void widthTest(float initialValue, float expectedValue) {
        Column column = new Column(0);
        assertEquals(Worksheet.DEFAULT_WORKSHEET_COLUMN_WIDTH, column.getWidth());
        column.setWidth(initialValue);
        assertEquals(expectedValue, column.getWidth());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing setWidth of Column")
    @ValueSource(
            floats = {
                    -1f,
                    255.1f}
    )
    void widthTest2(float value) {
        Column column = new Column(0);
        assertThrows(RangeException.class, () -> column.setWidth(value));
    }
}
